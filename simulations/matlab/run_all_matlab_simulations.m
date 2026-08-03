% =========================================================================
% run_all_matlab_simulations.m — Master Runner for All BMS Simulations
% MATLAB 2025b Compatible
% =========================================================================

clear; clc; close all;
set(0, 'DefaultFigureVisible', 'off');

fprintf('=====================================================\n');
fprintf('  Starting Self-Healing Cyber-Hardened BMS Sim Suite \n');
fprintf('=====================================================\n\n');

run('sim_ekf_adaptive.m');
fprintf('\n-----------------------------------------------------\n\n');
run('sim_ids_classifier.m');
fprintf('\n-----------------------------------------------------\n\n');
run('sim_5layer_bms_system.m');
fprintf('\n-----------------------------------------------------\n\n');
run('sim_7layer_self_healing_bms.m');

fprintf('\n=====================================================\n');
fprintf('  All 7-Layer Self-Healing BMS MATLAB Simulations OK!  \n');
fprintf('=====================================================\n');
