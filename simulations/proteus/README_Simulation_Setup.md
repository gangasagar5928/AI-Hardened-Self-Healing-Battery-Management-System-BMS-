# Master Simulation & 3D Circuit Guide — Cyber-Hardened BMS
**GCET EEE Senior Design Team | Academic Year 2025–2026**

This comprehensive guide covers how to execute all simulations across **Proteus 9 Pro**, **OrCAD X Pro+**, **LTSpice XVII**, **MATLAB R2025b**, and inspect the **3D Circuit Board** interactively.

---

## 1. 3D CIRCUIT BOARD SIMULATION & VISUALIZATION

### Interactive 3D WebGL Browser Viewer
We have provided a standalone interactive 3D WebGL viewer showing the complete 140mm x 100mm populated 2-layer FR4 printed circuit board.

- **File Location:** [`BMS_3D_Circuit_Viewer.html`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/BMS_3D_Circuit_Viewer.html)
- **How to Open:** Double-click `BMS_3D_Circuit_Viewer.html` or drag it into Chrome/Edge/Brave.
- **Features:**
  - **3D Orbit & Pan:** Left-click + drag to rotate 360°, Right-click + drag to pan, Scroll to zoom.
  - **Component Focus:** Click any component in the left panel (ESP32 U1, BQ76920 U2, SN65HVD230 U3, AMS1117 U4, OLED U5, 16S Header J1) to smoothly focus camera directly on that chip!
  - **Live Component Specs:** Shows footprint, signal speeds, and BOM pricing.

### KiCad 3D PCB Layout File
- **File Location:** [`simulations/kicad/Cyber_Hardened_BMS_3D.kicad_pcb`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/kicad/Cyber_Hardened_BMS_3D.kicad_pcb)
- **How to View in KiCad 7 / 8 / 10:**
  1. Open KiCad → File → Open Project → `simulations/kicad/Cyber_Hardened_BMS.kicad_pro`
  2. Click **PCB Editor** to open `Cyber_Hardened_BMS_3D.kicad_pcb`
  3. Press **Alt + 3** (or menu View → 3D Viewer) to open KiCad's built-in raytraced 3D Viewer!

---

## 2. PROTEUS 9 PRO SIMULATION GUIDE

### Files Provided:
- [`simulations/proteus/Cyber_Hardened_BMS.pdsprj`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/proteus/Cyber_Hardened_BMS.pdsprj) (Proteus 9 Design Project)
- [`simulations/proteus/BMS_Battery_Model.net`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/proteus/BMS_Battery_Model.net) (SPICE Netlist)

### Step-by-Step Instructions:
1. **Launch Proteus 9 Professional**.
2. **Open Project**: File → Open Project → Select `Cyber_Hardened_BMS.pdsprj`.
3. **Import SPICE Netlist**:
   - Go to top menu: `Design` → `Netlist Compiler`
   - Select Format: **SPICE Netlist**
   - Browse to `simulations/proteus/BMS_Battery_Model.net`
   - Click `OK` to load the 2RC Thevenin battery model and UDDS load current profile.
4. **Connect Virtual Oscilloscope**:
   - Select **Virtual Instruments** from left sidebar → **Oscilloscope**.
   - Connect Probe A to `CAN_H`, Probe B to `CAN_L`, Probe C to `N_SHUNT_OUT`, Probe D to `VCC_3V3`.
5. **Run VSM Simulation**:
   - Click the green **Play (▶)** button at bottom-left corner.
   - Observe battery terminal voltage `V(N_VOUT)` dropping dynamically during acceleration load spikes (UDDS drive profile).
6. **Simulate Cyber Attack**:
   - Open `BMS_Battery_Model.net` in text editor and uncomment line: `V_ATTACK N_VOUT 0 DC 12.0`.
   - Re-run simulation in Proteus → observe instant voltage spoofing collapse to 12.0V while actual current remains constant!

---

## 3. ORCAD X PRO+ / PSPICE SIMULATION GUIDE

### Files Provided:
- [`simulations/orcad/Cyber_Hardened_BMS.opj`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/orcad/Cyber_Hardened_BMS.opj) (OrCAD Express Project)
- [`simulations/orcad/Cyber_Hardened_BMS-PSpiceFiles/SCHEMATIC1/transient.sim`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/orcad/Cyber_Hardened_BMS-PSpiceFiles/SCHEMATIC1/transient.sim) (PSpice Simulation Profile)
- [`simulations/proteus/BMS_OrCAD_PSpice.cir`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/proteus/BMS_OrCAD_PSpice.cir) (PSpice Model Library)

### Step-by-Step Instructions:
1. **Launch OrCAD Capture CIS 24.1 / OrCAD X Pro+**.
2. **Open Project**: File → Open → Project → Select `simulations/orcad/Cyber_Hardened_BMS.opj`.
3. **Configure Simulation Profile**:
   - In Capture tree: `PSpice Resources` → `Simulation Profiles` → Right-click `transient` → Edit Profile.
   - Analysis Type: **Time Domain (Transient)**.
   - Run to Time: **30 seconds**, Start saving data after: **0s**, Transient Options: **0.01s**.
   - Under `Model Libraries` tab → Add `BMS_OrCAD_PSpice.cir` as Global Model.
4. **Execute PSpice Simulation**:
   - Click **PSpice** menu → **Run** (or press F11).
   - PSpice A/D Simulation Window opens.
5. **Plot Waveforms**:
   - Click `Trace` → `Add Trace` (or Insert key):
     - `V(N_3V3)`: 3.3V LDO rail stability under ESP32 100mA load.
     - `V(N_CANH) - V(N_CANL)`: Differential CAN bus voltage swing (0V recessive / 2V dominant at 500 kbps).
     - `V(N_TEMP1)`: NTC thermistor voltage divider output (~1.65V at 25°C).
     - `V(N_SHUNT_OUT)`: 10mΩ current sense voltage (1A = 10mV).

---

## 4. LTSPICE XVII SIMULATION GUIDE

### Files Provided in `simulations/ltspice/`:
1. [`cell_balancing_full.asc`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/ltspice/cell_balancing_full.asc) — 4S passive cell balancing equalisation waveforms.
2. [`can_bus_termination.asc`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/ltspice/can_bus_termination.asc) — CAN_H/CAN_L 120Ω split termination signal integrity at 500 kbps.
3. [`power_supply_3v3.asc`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/ltspice/power_supply_3v3.asc) — AMS1117-3.3 LDO regulator output ripple & dropout voltage.
4. [`ekf_rc_battery_model.asc`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/ltspice/ekf_rc_battery_model.asc) — 2RC Thevenin battery model during UDDS drive cycle.

### Step-by-Step Instructions:
1. Launch LTSpice XVII.
2. File → Open → Select any `.asc` file above.
3. Click green **Run (▶)** button or press **F5**.
4. Click on circuit nodes to probe voltage or components to probe current.

---

## 5. MATLAB R2025b SIMULATION GUIDE

### Files Provided in `simulations/matlab/`:
1. [`sim_7layer_self_healing_bms.m`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/matlab/sim_7layer_self_healing_bms.m) — Full 7-layer system Monte Carlo simulation across 50,000 frames.
2. [`sim_ids_classifier.m`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/matlab/sim_ids_classifier.m) — Trains Random Forest AI model & outputs confusion matrix graphic [`ids_confusion_matrix.png`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/matlab/ids_confusion_matrix.png).
3. [`sim_ekf_adaptive.m`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/matlab/sim_ekf_adaptive.m) — Evaluates Adaptive EKF SoC estimation under voltage spoofing & outputs graphic [`ekf_simulation_results.png`](file:///C:/Users/mksin/Desktop/AI%20hardened%20BMS/simulations/matlab/ekf_simulation_results.png).

---

## Hardware Bench Status Notice
> 📋 **Hardware Bench Testing Marked "To Be Tested / Pending"**:
> Physical HIL oscilloscope signal profiling, bench power analyzer measurement,
> and physical cell balancing trials are explicitly marked as "To Be Tested"
> for upcoming physical laboratory trials.
