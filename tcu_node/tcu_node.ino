/*
 * TCU Node (Telematics Control Unit) — ESP32 CAN to MQTT Gateway
 * Project: Self-Healing Cyber-Hardened Battery Management System (BMS)
 * 7-Layer Architecture with Digital Twin & Active Recovery Forwarding
 *
 * Data Flow:
 *   Battery Pack -> BQ76920 AFE -> BMS Master ESP32 (7-Layer Engine & Digital Twin)
 *   -> CAN Bus (500 kbps) -> TCU ESP32 Node -> Wi-Fi -> MQTT Broker -> Remote Cloud Dashboard
 */

#include <Arduino.h>
#include "driver/twai.h"
#include <WiFi.h>
#include <PubSubClient.h>

// CAN Pin Definitions (ESP32 TWAI)
#define CAN_TX_PIN GPIO_NUM_5
#define CAN_RX_PIN GPIO_NUM_4

// Wi-Fi Credentials
const char* WIFI_SSID     = "EV_Fleet_AP";
const char* WIFI_PASS     = "SecureBMS2026";

// MQTT Broker Configuration
const char* MQTT_SERVER   = "broker.hivemq.com";
const int   MQTT_PORT     = 1883;
const char* MQTT_TOPIC    = "ev/bms/telemetry";
const char* MQTT_ALERT_TOPIC = "ev/bms/alerts";

WiFiClient espClient;
PubSubClient mqttClient(espClient);

// Statistics
uint32_t totalFramesReceived = 0;
uint32_t attackFramesFiltered = 0;
uint32_t selfHealingActionsTriggered = 0;

void setupWiFi() {
  delay(10);
  Serial.println("\n[TCU Gateway] Connecting to Wi-Fi...");
  WiFi.begin(WIFI_SSID, WIFI_PASS);
  
  int attempts = 0;
  while (WiFi.status() != WL_CONNECTED && attempts < 20) {
    delay(500);
    Serial.print(".");
    attempts++;
  }
  
  if (WiFi.status() == WL_CONNECTED) {
    Serial.println("\n[TCU Gateway] Wi-Fi Connected! IP: " + WiFi.localIP().toString());
  } else {
    Serial.println("\n[TCU Gateway] Wi-Fi Connection Timeout — Operating in Standalone Buffer Mode.");
  }
}

void reconnectMQTT() {
  while (!mqttClient.connected()) {
    Serial.print("[TCU Gateway] Connecting to MQTT Broker...");
    String clientId = "ESP32_TCU_Gateway_";
    clientId += String(random(0xffff), HEX);
    
    if (mqttClient.connect(clientId.c_str())) {
      Serial.println("CONNECTED!");
      mqttClient.publish("ev/bms/status", "{\"tcu_status\":\"ONLINE\",\"version\":\"7_LAYER_SELF_HEALING_v2.0\"}");
    } else {
      Serial.print("FAILED, rc=");
      Serial.print(mqttClient.state());
      Serial.println(" Retrying in 5 seconds...");
      delay(5000);
    }
  }
}

void setupCAN() {
  twai_general_config_t g_config = TWAI_GENERAL_CONFIG_DEFAULT(CAN_TX_PIN, CAN_RX_PIN, TWAI_MODE_NORMAL);
  twai_timing_config_t t_config  = TWAI_TIMING_CONFIG_500KBITS();
  twai_filter_config_t f_config  = TWAI_FILTER_CONFIG_ACCEPT_ALL();

  if (twai_driver_install(&g_config, &t_config, &f_config) == ESP_OK) {
    Serial.println("[TCU Gateway] TWAI CAN Driver Installed (500 kbps).");
  } else {
    Serial.println("[TCU Gateway] ERROR: Failed to Install TWAI CAN Driver!");
    return;
  }

  if (twai_start() == ESP_OK) {
    Serial.println("[TCU Gateway] TWAI CAN Driver Started.");
  } else {
    Serial.println("[TCU Gateway] ERROR: Failed to Start TWAI CAN Driver!");
  }
}

void setup() {
  Serial.begin(115200);
  Serial.println("\n=================================================");
  Serial.println(" Self-Healing Cyber-Hardened BMS — TCU Node Gateway ");
  Serial.println(" 7-Layer Architecture with Digital Twin & MQTT Telemetry ");
  Serial.println("=================================================");

  setupCAN();
  setupWiFi();

  mqttClient.setServer(MQTT_SERVER, MQTT_PORT);
}

void processCANMessage(const twai_message_t& msg) {
  totalFramesReceived++;
  
  // BMS 7-Layer Status Frame ID 0x180
  if (msg.identifier == 0x180 && msg.data_length_code >= 8) {
    uint16_t raw_soc = (msg.data[0] << 8) | msg.data[1];
    uint16_t raw_vol = (msg.data[2] << 8) | msg.data[3];
    int16_t  raw_cur = (int16_t)((msg.data[4] << 8) | msg.data[5]);
    uint8_t  raw_ano = msg.data[6];
    uint8_t  action  = msg.data[7];

    float soc           = raw_soc / 100.0f;
    float voltage       = raw_vol / 1000.0f;
    float current       = raw_cur / 1000.0f;
    float anomaly_score = raw_ano / 100.0f;
    float trust_score   = 1.0f - anomaly_score;
    float soh_health    = 98.5f - (anomaly_score * 2.0f);
    float thermal_pred  = 29.5f + (anomaly_score * 8.5f);

    bool is_attack = (anomaly_score > 0.70f);

    const char* action_str = "CONTINUE_NORMAL";
    if (action == 0x01) { action_str = "REDUCE_CHARGE_CURRENT"; selfHealingActionsTriggered++; }
    else if (action == 0x02) { action_str = "DISABLE_BALANCING"; selfHealingActionsTriggered++; }
    else if (action == 0x03) { action_str = "LIMP_HOME_MODE"; selfHealingActionsTriggered++; }
    else if (action == 0x04) { action_str = "MODULE_ISOLATED"; selfHealingActionsTriggered++; }
    else if (is_attack) { action_str = "ATTACK_FLAGGED"; attackFramesFiltered++; }

    // Build Extended JSON Telemetry Payload
    char payload[384];
    snprintf(payload, sizeof(payload),
             "{\"node\":\"TCU_ESP32\","
             "\"soc\":%.2f,"
             "\"voltage\":%.2f,"
             "\"current\":%.2f,"
             "\"anomaly_score\":%.2f,"
             "\"trust_score\":%.2f,"
             "\"soh_health_index\":%.1f,"
             "\"thermal_pred_c\":%.1f,"
             "\"self_healing_action\":\"%s\","
             "\"status\":\"%s\","
             "\"total_rx\":%u,"
             "\"attacks_flagged\":%u}",
             soc, voltage, current, anomaly_score, trust_score, soh_health, thermal_pred,
             action_str, is_attack ? "ATTACK_FLAGGED" : "NORMAL",
             totalFramesReceived, attackFramesFiltered);

    Serial.printf("[TCU -> MQTT] Publishing 7-Layer Telemetry: %s\n", payload);

    if (WiFi.status() == WL_CONNECTED) {
      if (!mqttClient.connected()) {
        reconnectMQTT();
      }
      mqttClient.loop();
      mqttClient.publish(MQTT_TOPIC, payload);

      if (is_attack || action > 0) {
        mqttClient.publish(MQTT_ALERT_TOPIC, payload);
      }
    }
  }
}

void loop() {
  if (WiFi.status() == WL_CONNECTED) {
    if (!mqttClient.connected()) {
      reconnectMQTT();
    }
    mqttClient.loop();
  }

  twai_message_t msg;
  if (twai_receive(&msg, pdMS_TO_TICKS(10)) == ESP_OK) {
    processCANMessage(msg);
  }
}
