# Cyber-Hardened BMS — Step-by-Step Build & Setup Guide

This guide provides step-by-step instructions for assembling hardware, setting up the software development environment, compiling firmware, training the TinyML classifier, and executing Hardware-in-the-Loop (HIL) validation tests.

---

## 1. Prerequisites & Software Toolchain Setup

### Hardware Requirements
* 2x ESP32 WROOM-32D Development Boards (or 1x ESP32-S3 / bare ICs for cost optimization).
* 1x TI BQ76920 AFE breakout module or custom PCB.
* 2x SN65HVD230 / VP230 3.3V CAN Transceiver modules.
* 1x High-Side P-FET / Optocoupler Solid-State Relay (SSR) driver module.
* 4x 18650 High-Drain Li-ion Cells (4S1P configuration, $14.8\text{V}$ nominal).
* 4x IRLML2502 Logic-Level N-Channel MOSFETs + 4x $47\,\Omega$ ($1\text{W}$) resistors.
* 1x MicroSD SPI Card Module + 16GB MicroSD card (FAT32 formatted).

---

### Software Installation

1. **Install Arduino IDE 2.x:**
   Download and install from [arduino.cc](https://www.arduino.cc/en/software).

2. **Add ESP32 Board Package (v3.0+):**
   In Arduino IDE preferences, add URL:
   `https://espressif.github.io/arduino-esp32/package_esp32_index.json`
   Open Board Manager, search for `esp32` by Espressif Systems, and click **Install**.

3. **Install Required Arduino Libraries:**
   * `Wire` (Built-in)
   * `SPI` (Built-in)
   * `Preferences` (Built-in)
   * `Adafruit_SSD1306` & `Adafruit_GFX`

4. **Install Python 3.10+ Dependencies:**
   ```bash
   pip install numpy pandas scikit-learn matplotlib mbedtls m2cgen pyserial
   ```

---

## 2. Low-Cost Optimized BOM Purchasing Reference

Target Total Budget: **₹2,100 – ₹2,400 per setup**

| Component | Source / Alternative | Recommended Spec | Unit Price |
| :--- | :--- | :--- | :--- |
| **Microcontrollers** | ESP32-WROOM-32D / ESP32-S3 | 2x NodeMCU ESP32 Boards | ₹345 / unit |
| **CAN Transceivers** | VP230 / TJA1051 ICs | Standalone DIP/SOP Transceiver ICs | ₹60 / unit |
| **Battery Pack** | 4S1P Reclaimed Pack | High-Drain 18650 Cells (Samsung 25R equivalent) | ₹110 / cell |
| **SSR Cutoff Switch** | Optocoupler PC817 + IRF4905 | High-Side P-Channel FET Cutoff Driver | ₹90 |
| **AFE Board** | TI BQ76920 Breakout | 3S-5S Cell Monitoring Breakout | ₹800 |

---

## 3. Hardware Assembly Instructions

1. **Power Bus Setup:** Connect common ground (`GND`) across Master ESP32, Attacker ESP32, BQ76920 AFE, and CAN Transceivers.
2. **CAN Bus Setup:**
   * Master ESP32 GPIO 5 (`TX`) $\to$ SN65HVD230 #1 `TXD`.
   * Master ESP32 GPIO 4 (`RX`) $\to$ SN65HVD230 #1 `RXD`.
   * Attacker ESP32 GPIO 5 (`TX`) $\to$ SN65HVD230 #2 `TXD`.
   * Attacker ESP32 GPIO 4 (`RX`) $\to$ SN65HVD230 #2 `RXD`.
   * Connect `CANH` to `CANH`, `CANL` to `CANL`. Connect $120\,\Omega$ split termination resistors at both bus ends.
3. **AFE Sensor Wiring:**
   * Connect 4S cell balance leads to VC1, VC2, VC3, VC4 pins on BQ76920.
   * Connect I2C `SDA` (GPIO 21) and `SCL` (GPIO 22) with $4.7\text{ k}\Omega$ pullup resistors to 3.3V.
   * Connect `ALERT` interrupt pin to GPIO 35.
4. **High-Side SSR Cutoff Wiring:**
   * Connect Master ESP32 GPIO 17 to optocoupler input resistor ($220\,\Omega$).
   * Connect optocoupler output to P-FET gate pull-down network on main battery positive line.

---

## 4. Firmware Compilation & Flashing

### Flash Master Node (`bms_master`)
1. Connect Master ESP32 via USB.
2. Open `bms_master/bms_master.ino` in Arduino IDE.
3. Select Board: `ESP32 Dev Module`, CPU Frequency: `240MHz`, Partition Scheme: `Huge APP (3MB)`.
4. Click **Upload**.

### Flash Attacker Node (`attacker_node`)
1. Connect Attacker ESP32 via USB.
2. Open `attacker_node/attacker_node.ino`.
3. Set `ATTACK_MODE = 5;` for mixed attack generation.
4. Click **Upload**.

---

## 5. Machine Learning Dataset Capture & Training

### 1. Capture CAN Telemetry
Connect Attacker ESP32 serial port (e.g. `COM3` / `/dev/ttyUSB0`) and run:
```bash
python generate_dataset.py COM3
```
Let it record ~10,000 frames under clean and injected traffic. Press `Ctrl+C` to save `can_dataset.csv`.

### 2. Train Random Forest & Export C++ Tree
```bash
python train_ids.py
```
This trains a Random Forest model on features ($\Delta t, f, \text{Var}, H$) and compiles it into `ids_model.h` using `m2cgen`. Recompile `bms_master.ino` to deploy the updated weights.

---

## 6. Verification & HIL Testing Protocol

### 1. Run Master Simulation Suite
```bash
python run_all_simulations.py
```
Verify generated plots in `simulations/matlab/`:
* `sim_7layer_results.png`
* `sim_ekf_adaptive_results.png` (SoC tracking error $<1.4\%$)
* `sim_ids_classifier_results.png` (6-class accuracy $>99.2\%$)

### 2. Oscilloscope Latency Verification
* Connect scope Channel 1 to CAN `TX` pin.
* Connect scope Channel 2 to GPIO 17 (`SSR_CUTOFF`).
* Trigger an emergency injection frame (Attack Mode 7).
* Verify propagation delay from frame reception to GPIO 17 HIGH is **$< 1.2\text{ ms}$**.
