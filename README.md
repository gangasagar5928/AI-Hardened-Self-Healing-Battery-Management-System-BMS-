<p align="center">
  <img src="assets/logo.png" width="480" alt="Cyber-Hardened Battery Management System (BMS) Logo">
</p>

<h1 align="center">Cyber-Hardened Battery Management System (BMS) for EVs</h1>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-blue.svg" alt="License: MIT"></a>
  <a href="https://www.espressif.com/en/products/socs/esp32"><img src="https://img.shields.io/badge/Platform-ESP32%20FreeRTOS-orange.svg" alt="Platform: ESP32"></a>
  <a href="#7-layer-security-architecture"><img src="https://img.shields.io/badge/Architecture-7--Layer%20Hardened-green.svg" alt="Architecture"></a>
  <a href="#tinyml--algorithmic-innovation"><img src="https://img.shields.io/badge/TinyML-m2cgen%20C%2B%2B%20%3C0.35ms-red.svg" alt="TinyML"></a>
</p>

An open-source, production-grade **Self-Healing Cyber-Hardened Battery Management System** designed for Electric Vehicle (EV) 2-wheeler and 3-wheeler applications. The system protects unauthenticated CAN and Bluetooth (BLE) buses against Denial of Service (DoS) floods, false data injection (spoofing), replay attacks, and UDS (ISO 14229) session hijacks.

It introduces a novel **Trust-Scaled Extended Kalman Filter (EKF)** that dynamically couples an on-device machine learning anomaly score $S \in [0, 1]$ to scale the measurement covariance matrix $R_{\text{eff}} = R_{\text{base}} \cdot e^{\lambda \cdot S}$. Under cyber-attack ($S \to 1$), the Kalman gain approaches zero ($K \to 0$), suppressing corrupted sensor data without controller collapse.

---

## Key Highlights & High-Impact Novelty

* **Zero-GPU Embedded Execution:** Compiled via `m2cgen` into C++ decision tree `if-else` blocks running in `<0.35 ms` latency on ESP32 Core 0 ($38.4\text{ KB}$ SRAM footprint).
* **Dual-Core FreeRTOS Isolation:** Core 0 handles ML anomaly classification, BLE auth, and UDS inspection; Core 1 executes deterministic AFE polling, EKF state estimation, cell balancing, and high-side SSR cutoff control.
* **Automated High-Side SSR Cutoff:** Automated physical pack disconnect triggered on GPIO 17 within $<1.2\text{ ms}$ during emergency severe attacks ($S > 0.90$).
* **Layer 3 UDS (ISO 14229) Protection:** Intercepts unauthorized `SecurityAccess` (`0x27`) seed requests and `TesterPresent` (`0x3E`) keep-alives over OBD-II/CAN IDs `0x7E0`/`0x7E8`.
* **📄 Digital Battery Health Passport:** Cryptographically signed (SHA-256 digest) tamper-evident battery history certificate (SoH curve, cycle count, thermal spikes) for used-EV resale trust infrastructure.
* **🔒 Battery Theft & Physical Tamper Detection:** Extended Layer 3 TinyML IDS detecting abrupt pack removal and physical cell-bypassing jumpering.
* **⚡ Grid-Aware Adaptive Demand-Response Charging:** Pure firmware scheduling logic throttling peak grid charging by 50% during grid stress signals (CAN `0x198`).
* **🤖 Federated Learning Fleet Intelligence:** Privacy-preserving model-weight delta sharing (`0x188` CAN ID) allowing deployed BMS units to locally fine-tune and share updates across the fleet.
* **Cost-Optimized BOM:** Reduced total setup cost to **₹2,100 – ₹2,400 (~$28–$32 USD)** per prototype benchmark setup.

---

## 7-Layer Security Architecture

```mermaid
graph TD
    A["CAN Bus / BLE Telemetry"] --> B["Layer 1: ECDH + AES-128 BLE Auth"]
    B --> C["Layer 2: HMAC Command Verification"]
    C --> D["Layer 3: TinyML IDS & UDS Inspector (0x7E0)"]
    D --> E["Layer 4: Adaptive EKF Covariance Scaling R_eff = R_base * e^(10*S)"]
    E --> F["Layer 5: Self-Healing Decision Engine"]
    F --> G{"Anomaly Score S?"}
    G -- "S < 0.30" --> H["Normal Operation"]
    G -- "0.30 <= S < 0.50" --> I["Reduce Charge Current"]
    G -- "0.50 <= S < 0.70" --> J["Disable Cell Balancing"]
    G -- "0.70 <= S < 0.90" --> K["Limp Home Mode"]
    G -- "S >= 0.90" --> L["High-Side SSR Pack Isolation (GPIO 17 HIGH)"]
    F --> M["Layer 6: Cell Reconfiguration & Passive Balancing"]
    F --> N["Layer 7: Secure TCU Cloud Gateway (MQTT / Node-RED)"]
```

---

## Repository Structure

```
├── assets/                   # Repository logo & visual media assets
│   └── logo.png
├── bms_master/               # ESP32 Master Firmware (FreeRTOS Core 0/1, EKF, IDS, SSR driver)
│   ├── bms_master.ino
│   └── ids_model.h           # Compiled m2cgen C++ decision tree weights
├── attacker_node/            # ESP32 Attack Injector (Modes 1-7: DoS, Spoof, Replay, Fuzz, UDS, SSR Test)
│   └── attacker_node.ino
├── tcu_node/                 # Telematics Control Unit MQTT bridge
│   └── tcu_node.ino
├── docs/                     # Full GitHub Documentation Suite (Wiki, API, Wiring, Flowcharts, Build Guide)
│   ├── WIKI.md
│   ├── API_DOCUMENTATION.md
│   ├── FLOWCHARTS.md
│   ├── WIRING_DIAGRAMS.md
│   └── BUILD_GUIDE.md
├── simulations/              # MATLAB, KiCad, SPICE & Proteus simulation suite
│   ├── matlab/               # .m scripts & generated PNG plot waveforms
│   └── kicad/ ltspice/ ...   # Hardware schematics & spice files
├── generate_dataset.py       # Live CAN telemetry & feature extraction script
├── train_ids.py              # Random Forest training & C++ code export
└── run_all_simulations.py    # Master python simulation & visualization runner
```

---

## Documentation Quick Links

* 📖 **[Project Wiki](docs/WIKI.md)** — In-depth theory, EKF math derivation, 4S vs 16S pack architecture, and secure boot setup.
* 🔌 **[API Documentation](docs/API_DOCUMENTATION.md)** — CAN message IDs (`0x180`, `0x7E0`), UDS diagnostic services, BLE UUIDs, and firmware APIs.
* 📊 **[Flowcharts & Diagrams](docs/FLOWCHARTS.md)** — Interactive Mermaid diagrams for FreeRTOS tasks, EKF trust scaling, and self-healing.
* ⚡ **[Wiring Diagrams](docs/WIRING_DIAGRAMS.md)** — Pinout reference, High-Side SSR cutoff schematics, and BQ76920 AFE connections.
* 🛠️ **[Build & Setup Guide](docs/BUILD_GUIDE.md)** — Step-by-step assembly, low-cost BOM purchasing table, code compilation, and HIL testing.

---

## Low-Cost Bill of Materials (Target ₹2,100 – ₹2,400)

| Component Category | Baseline Choice | Low-Cost Optimized Choice | Savings | Optimized Price |
| :--- | :--- | :--- | :--- | :--- |
| **Microcontrollers** | 3x ESP32 Dev Boards (₹1,140) | 1x ESP32-S3 / 2x Bare ICs + Laptop SLCAN | ₹380 – ₹450 | **₹690** |
| **CAN Transceivers** | 3x SN65HVD230 Modules (₹360) | 2x VP230 / TJA1051 Standalone DIP/SOP ICs | ₹120 | **₹240** |
| **Battery Pack** | 16x New 18650 Cells (₹1,050) | 4S1P Reclaimed High-Drain 18650 Pack | ₹600 | **₹450** |
| **Displays** | 0.96" SSD1306 OLED (₹90) | Stream to Node-RED / Grafana / Serial | ₹90 | **₹0** |
| **AFE & Sensors** | BQ76920 Breakout Board | BQ76920 IC + Shunt Resistor | ₹0 | **₹800** |
| **High-Side SSR Cutoff** | Standard Relay (₹120) | High-Side SSR / Optocoupler P-FET Switch | ₹30 | **₹90** |
| **Total Setup** | **₹3,600 – ₹4,750** | **Target Prototype Bench Setup** | **₹1,220 – ₹1,290** | **₹2,100 – ₹2,400** |

---

## Quick Start Guide

Follow these 5 steps to run simulations, train the TinyML model, flash firmware nodes, and verify hardware performance.

### 1. Install Dependencies
```bash
# Clone repository
git clone https://github.com/gangasagar5928/AI-Hardened-Self-Healing-Battery-Management-System-BMS-.git
cd AI-Hardened-Self-Healing-Battery-Management-System-BMS-

# Install Python requirements
pip install numpy pandas scikit-learn matplotlib m2cgen pyserial
```

### 2. Run Master Simulation Suite
Execute the 5-part simulation suite (Monte Carlo, EKF tracking, TinyML classifier, SPICE waveforms, Digital Twin):
```bash
python run_all_simulations.py
```
Generates 5 high-resolution plot artifacts in `simulations/matlab/`:
* `sim_7layer_results.png` (7-Layer system resilience & self-healing counts)
* `sim_ekf_adaptive_results.png` (Adaptive EKF vs Standard EKF SoC tracking under attack)
* `sim_ids_classifier_results.png` (6-class TinyML confusion matrix)
* `sim_spice_circuits_results.png` (CAN bus split termination & MOSFET balancing waveforms)
* `sim_digital_twin_thermal_results.png` (Predictive thermal & SoH degradation trajectory)

### 3. Capture Dataset & Train TinyML Classifier
1. Connect Attacker ESP32 to USB port (e.g. `COM3` on Windows or `/dev/ttyUSB0` on Linux) and log live telemetry:
   ```bash
   python generate_dataset.py COM3
   ```
2. Train the Random Forest classifier and compile weights into C++ header:
   ```bash
   python train_ids.py
   ```
   *Generates `bms_master/ids_model.h` containing C++ decision tree logic generated via `m2cgen`.*

### 4. Flashing Embedded Firmware Nodes

Open [Arduino IDE 2.x](https://www.arduino.cc/en/software), install the **ESP32 Board Package (v3.0+)**, select Board `ESP32 Dev Module`, Partition Scheme `Huge APP (3MB)`:

* **BMS Master Node (`bms_master/bms_master.ino`):**
  - Connect Master ESP32 via USB.
  - Flashes Core 0 TinyML IDS, Layer 3 UDS Inspector, Layer 4 Adaptive EKF, and GPIO 17 High-Side SSR cutoff driver.
  - Click **Upload**.

* **Attacker Node (`attacker_node/attacker_node.ino`):**
  - Connect Attacker ESP32 via USB.
  - Set `ATTACK_MODE = 5;` to cycle through all 7 attack vectors (DoS, Spoof, Replay, Fuzz, UDS Hijack, Emergency SSR Test).
  - Click **Upload**.

* **TCU Cloud Gateway (`tcu_node/tcu_node.ino`):**
  - Set Wi-Fi SSID/password and MQTT broker URL.
  - Click **Upload** to stream telemetry to Node-RED / Grafana dashboards.

### 5. Hardware-in-the-Loop (HIL) Verification
1. Connect CAN Bus: Master `CAN TX` (GPIO 5) / `CAN RX` (GPIO 4) to SN65HVD230 transceiver, wired to Attacker Node transceiver (`CANH` to `CANH`, `CANL` to `CANL`).
2. Open Serial Monitor at `115200` baud.
3. Observe real-time output:
   - Dynamic measurement noise covariance scaling ($R_{\text{eff}} = R_{\text{base}} \cdot e^{10 \cdot S}$).
   - Automated GPIO 17 High-Side SSR cutoff activation within $<1.2\text{ ms}$ when anomaly score $S > 0.90$.
   - SoC estimation error maintained below $<1.4\%$ during sustained CAN attacks.

---

## License

Distributed under the **MIT License**. See `LICENSE` for details.
