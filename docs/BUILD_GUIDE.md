# Cyber-Hardened BMS — Step-by-Step Build & Setup Guide (With Beginner Explanations)

This guide provides step-by-step instructions for assembling hardware, setting up the software development environment, compiling firmware, training the TinyML classifier, and executing Hardware-in-the-Loop (HIL) validation tests.

---

## 1. Beginner Guide — Understanding the 4 Advanced Features in Simple Terms

If you are new to battery management systems or cybersecurity, here is a simple, easy-to-understand breakdown of what these 4 new features do and why they matter:

### 📄 Feature 1: Digital Battery Health Passport
* **What is it?** Think of it like a digital, un-hackable "service record card" or "birth certificate" for the battery.
* **Why do we need it?** When buying a used EV or second-hand battery in India, buyers have no idea if the previous owner abused the battery or if it's about to die. Sellers can fake numbers.
* **How does it work?** The BMS continuously logs battery health, charge cycles, and high temperatures, and calculates a 256-bit cryptographic signature (SHA-256 hash). If anyone tries to alter the data, the signature breaks instantly. Buyers can scan a QR code to verify the real battery condition before buying!

### 🔒 Feature 2: Battery Theft & Physical Cell-Bypass Tamper Detection
* **What is it?** An automatic alarm inside the BMS that detects if someone is physically stealing the battery pack or removing individual cells.
* **Why do we need it?** Batteries account for 40%–50% of an EV's cost. Battery theft and cell harvesting are rising problems. Thieves sometimes bypass parts of the battery with jumper wires to steal cells.
* **How does it work?** The TinyML AI monitors sudden voltage drops and electrical jumps. If a thief unplugged the pack while current was flowing, or jumped a cell, the BMS logs a non-erasable theft alert and locks the main power switch (SSR cutoff).

### ⚡ Feature 3: Grid-Aware Adaptive Demand-Response Charging
* **What is it?** Smart charging logic that protects the city's electricity grid during peak hours.
* **Why do we need it?** If thousands of EVs plug in at 7:00 PM when everyone turns on ACs and lights, the power grid can overload and cause blackouts.
* **How does it work?** When the BMS detects peak grid stress, it automatically lowers its charging speed by 50% or postpones fast charging until off-peak hours (e.g. late night). It requires zero extra hardware—just smart software scheduling!

### 🤖 Feature 4: Federated Learning Fleet Intelligence
* **What is it?** A way for all BMS units on the road to learn about new cyber-attacks together without leaking private driver data.
* **Why do we need it?** Hackers constantly invent new attacks. A static AI model trained in a lab will miss new threats after a few months.
* **How does it work?** Instead of sending your personal location or battery logs to a central server, each BMS updates its AI model locally. It only shares small mathematical "lessons" (model weight deltas). The central server combines these lessons to upgrade the entire fleet automatically!

---

## 2. Prerequisites & Software Toolchain Setup

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
   pip install numpy pandas scikit-learn matplotlib m2cgen pyserial
   ```

---

## 3. Low-Cost Optimized BOM Purchasing Reference

Target Total Budget: **₹2,100 – ₹2,400 per setup**

| Component | Source / Alternative | Recommended Spec | Unit Price |
| :--- | :--- | :--- | :--- |
| **Microcontrollers** | ESP32-WROOM-32D / ESP32-S3 | 2x NodeMCU ESP32 Boards | ₹345 / unit |
| **CAN Transceivers** | VP230 / TJA1051 ICs | Standalone DIP/SOP Transceiver ICs | ₹60 / unit |
| **Battery Pack** | 4S1P Reclaimed Pack | High-Drain 18650 Cells (Samsung 25R equivalent) | ₹110 / cell |
| **SSR Cutoff Switch** | Optocoupler PC817 + IRF4905 | High-Side P-Channel FET Cutoff Driver | ₹90 |
| **AFE Board** | TI BQ76920 Breakout | 3S-5S Cell Monitoring Breakout | ₹800 |

---

## 4. Hardware Assembly Instructions

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

## 5. Firmware Compilation & Flashing

### Flash Master Node (`bms_master`)
1. Connect Master ESP32 via USB.
2. Open `bms_master/bms_master.ino` in Arduino IDE.
3. Select board `ESP32 Dev Module`, Partition Scheme `Huge APP (3MB)`.
4. Click **Upload**.

### Flash Attacker Node (`attacker_node`)
1. Connect Attacker ESP32 via USB.
2. Open `attacker_node/attacker_node.ino` in Arduino IDE.
3. Set `ATTACK_MODE = 5;` (Mixed 1–7 cycling).
4. Click **Upload**.

---

## 6. HIL Testing & Serial Verification

1. Connect both ESP32 boards over CAN bus.
2. Open Serial Monitor at `115200` baud.
3. Observe live output:
   - `[ATTACK]` notifications.
   - Dynamic covariance scaling $R_{\text{eff}} = R_{\text{base}} \cdot e^{10 \cdot S}$.
   - Automated GPIO 17 SSR Cutoff trip when anomaly score $S > 0.90$.
   - SHA-256 Digital Battery Passport verification digest logged over CAN ID `0x190`.
