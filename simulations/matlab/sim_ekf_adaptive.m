% =========================================================================
% sim_ekf_adaptive.m — Cyber-Hardened BMS: Adaptive EKF SoC Estimation
% MATLAB 2025b & All Releases Compatible (Toolbox Independent)
% =========================================================================

clear; clc; close all;

fprintf('=====================================================\n');
fprintf('  BMS Adaptive EKF Simulation (Dynamic Covariance)\n');
fprintf('=====================================================\n\n');

dt = 0.1;                       % Sampling interval (s)
t_total = 600;                  % Total simulation time (s)
time = 0:dt:t_total;
N = length(time);

Q_nom_Ah = 2.5;                 % Capacity (Ah)
Q_nom_As = Q_nom_Ah * 3600;     % Capacity (As)
R0 = 0.05;                      % Ohmic resistance (Ohms)
R1 = 0.02;                      % Polarization resistance (Ohms)
C1 = 1000;                      % Polarization capacitance (Farads)
tau = R1 * C1;                  % RC time constant (20 s)

R_base = 0.01;                  % Base measurement noise variance (V^2)
Q_proc = diag([1e-6, 1e-8]);    % Process noise covariance matrix

I_load = 2.5 * ones(1, N); 
I_load(1500:2000) = 4.0;        % Pulse load 1
I_load(3500:4000) = 1.0;        % Light load period

x_true = zeros(2, N);
x_true(:, 1) = [0.95; 0.0];

for k = 1:N-1
    i_curr = I_load(k);
    dSoC = -(i_curr * dt) / Q_nom_As;
    dVrc = -dt/tau * x_true(2, k) + (R1/tau) * i_curr * dt;
    x_true(1, k+1) = max(0, min(1, x_true(1, k) + dSoC));
    x_true(2, k+1) = x_true(2, k) + dVrc;
end

ocv_fit = @(soc) 3.0 + 1.15*soc + 0.2*soc.^2 - 0.15*soc.^3;
d_ocv_fit = @(soc) 1.15 + 0.4*soc - 0.45*soc.^2;

V_true = ocv_fit(x_true(1, :)) - x_true(2, :) - I_load * R0;

% Attack window: k = 2001 to 4000
attack_start_k = 2001;
attack_end_k = 4000;

V_meas = V_true + sqrt(R_base) * randn(1, N);
anomaly_score = zeros(1, N);

attack_bias = 0.35;
V_meas(attack_start_k:attack_end_k) = V_meas(attack_start_k:attack_end_k) + attack_bias;
anomaly_score(attack_start_k:attack_end_k) = 0.98;

% Moving average filter (Toolbox independent filter replacement)
b_filter = ones(1,10)/10;
anomaly_score = filter(b_filter, 1, anomaly_score);

x_std = zeros(2, N);
x_std(:, 1) = [0.90; 0.0];
P_std = diag([1e-3, 1e-4]);

x_adapt = zeros(2, N);
x_adapt(:, 1) = [0.90; 0.0];
P_adapt = diag([1e-3, 1e-4]);
R_eff_hist = zeros(1, N);
K_gain_hist = zeros(1, N);

for k = 1:N-1
    i_curr = I_load(k);
    
    % Standard EKF
    x_std_pred = [x_std(1,k) - (i_curr*dt)/Q_nom_As; ...
                  x_std(2,k)*exp(-dt/tau) + R1*(1-exp(-dt/tau))*i_curr];
    A = [1, 0; 0, exp(-dt/tau)];
    P_std_pred = A * P_std * A' + Q_proc;
    
    h_std = ocv_fit(x_std_pred(1)) - x_std_pred(2) - i_curr * R0;
    H_std = [d_ocv_fit(x_std_pred(1)), -1];
    K_std = P_std_pred * H_std' / (H_std * P_std_pred * H_std' + R_base);
    x_std(:, k+1) = x_std_pred + K_std * (V_meas(k+1) - h_std);
    P_std = (eye(2) - K_std * H_std) * P_std_pred;
    
    % Adaptive EKF
    x_adapt_pred = [x_adapt(1,k) - (i_curr*dt)/Q_nom_As; ...
                    x_adapt(2,k)*exp(-dt/tau) + R1*(1-exp(-dt/tau))*i_curr];
    P_adapt_pred = A * P_adapt * A' + Q_proc;
    
    S_anom = anomaly_score(k+1);
    R_eff = R_base * exp(10 * S_anom);
    R_eff_hist(k+1) = R_eff;
    
    h_adapt = ocv_fit(x_adapt_pred(1)) - x_adapt_pred(2) - i_curr * R0;
    H_adapt = [d_ocv_fit(x_adapt_pred(1)), -1];
    K_adapt = P_adapt_pred * H_adapt' / (H_adapt * P_adapt_pred * H_adapt' + R_eff);
    K_gain_hist(k+1) = K_adapt(1);
    
    x_adapt(:, k+1) = x_adapt_pred + K_adapt * (V_meas(k+1) - h_adapt);
    P_adapt = (eye(2) - K_adapt * H_adapt) * P_adapt_pred;
end

fig = figure('Position', [100, 100, 1000, 800], 'Color', 'w');

subplot(4, 1, 1);
plot(time, V_true, 'k-', 'LineWidth', 1.5); hold on;
plot(time, V_meas, 'r.', 'MarkerSize', 4);
ylabel('Voltage (V)');
title('Terminal Voltage & False Data Injection Attack (4S Modular Sub-Unit, 14.8V Nominal)');
legend('True Voltage', 'Measured (Spoofed)', 'Location', 'SouthWest');
grid on;

subplot(4, 1, 2);
plot(time, x_true(1,:)*100, 'k-', 'LineWidth', 2); hold on;
plot(time, x_std(1,:)*100, 'r--', 'LineWidth', 1.5);
plot(time, x_adapt(1,:)*100, 'b-', 'LineWidth', 1.5);
ylabel('SoC (%)');
title('State-of-Charge Estimation Comparison (Adaptive EKF Hardening)');
legend('True SoC', 'Standard EKF (Vulnerable)', 'Adaptive EKF (Cyber-Hardened)', 'Location', 'SouthWest');
grid on;

subplot(4, 1, 3);
err_std = abs(x_true(1,:) - x_std(1,:)) * 100;
err_adapt = abs(x_true(1,:) - x_adapt(1,:)) * 100;
plot(time, err_std, 'r-', 'LineWidth', 1.2); hold on;
plot(time, err_adapt, 'b-', 'LineWidth', 1.5);
ylabel('Error (%)');
title('Absolute State-of-Charge Error (%)');
legend('Standard EKF Error', 'Adaptive EKF Error', 'Location', 'NorthWest');
grid on;

subplot(4, 1, 4);
yyaxis left
plot(time, anomaly_score, 'm-', 'LineWidth', 1.5); hold on;
ssr_trip = (anomaly_score > 0.90) * 1.0;
plot(time, ssr_trip, 'r:', 'LineWidth', 2.0);
ylabel('Anomaly Score S & SSR Cutoff (1=Trip)');
yyaxis right
semilogy(time, R_eff_hist, 'g-', 'LineWidth', 1.5);
ylabel('R_{eff} (V^2)');
xlabel('Time (seconds)');
title('Layer 3 IDS Anomaly Score, High-Side SSR Cutoff & Layer 4 Dynamic Covariance R_{eff}');
grid on;

saveas(fig, 'ekf_simulation_results.png');
fprintf('Simulation complete. Plot saved as ekf_simulation_results.png.\n');
fprintf('Max Standard EKF SoC Error: %.2f%%\n', max(err_std));
fprintf('Max Adaptive EKF SoC Error: %.2f%%\n', max(err_adapt));
