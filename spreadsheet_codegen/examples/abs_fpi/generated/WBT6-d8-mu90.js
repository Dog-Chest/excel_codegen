// WBT6-d8-mu90
// See ext_body.js.j2 for why the coordinates go through temporaries.
var t1 = x;
var t2 = y;
var t3 = z;
x = t1;
y = t2;
z = t3;
// --- global parameters (same for every case)
var L = 340 m;
var B = 64 m;
var D = 32 m;
var rho_sea = 1025 kg/m^3;
var g = 9.81 m/s^2;
var Cb = 1;
var d_f = 20.559 m;
var rho_g_lo = 10050 N/m^3;
var rho_g_hi = 11180 N/m^3;
// --- this case
var draft = 8 m;
var mu_deg = 90 deg;
var k_c = 1;
var k_u = 1.1;
var beta_VAC = 0.66;
var beta_LAC = 0.758;
var beta_TAC = 0.557;
var beta_PMO = 0.686;
var beta_RMO = 0.492;
var k_esf = 1.1;
var w_v = 0.75;
var w_l = 0.25;
var w_t = 0.75;
var xi = 21 m;
var delta_b = 0 m;
var delta_h = 0 m;
// --- the tank this case names: WBT6
var rho_tank = 1025 kg/m^3;
var l_tank = 42 m;
var b_tank = 32 m;
var h_tank = 32 m;
var eta_deck = 0 m;
var eta_overflow = 0 m;
var C_dp = 1;
var C_ru = 1;
var p_vp = 0 N/cm^2;
var GM_full_in = 0 m;
var k_r_in = 0 m;
var tank_is_ballast = 1;
var member_11_17 = 0;
// --- the calculation
var C_theta = 1;   // C_theta, the roll angle factor of 5.7.2(b)
var C_phi = -0.45;   // C_phi, the pitch angle factor of 5.7.2(b)
var k_theta = 0.005;   // k_theta of the theta expression (5.7.1(b))
var C_R = 1.05;   // C_R of the theta expression (5.7.1(b))
var k_o = 1.34 - 0.47 * Cb;   // k_o (5.7.1(c))
var a_o = (k_o * (2.4 / Math.pow(L / 1 m, 0.5) + 34 / (L / 1 m) - 600 / Math.pow(L / 1 m, 2))) * 1 m;   // a_o, the vertical acceleration parameter
var D_1 = (10 / Cb);   // the (10 / C_b) term of phi
var phi = beta_PMO * k_esf * 1030 * Math.pow(D_1, 0.25) / (L / 1 m) * 1 deg;   // phi, the pitch amplitude
if (phi > 10 deg) phi = 10 deg;
else              phi = phi;
var GM_full = GM_full_in > 0 m ? GM_full_in : 0.12 * B;   // 5.7.1(b): 0.12 B when GM(full) is not available
var k_r = k_r_in > 0 m ? k_r_in : (draft <= d_f ? 0.35 * B : 0.45 * B);   // roll radius of gyration
var Delta = 10.05 kN/m^3 * L * B * d_f * Cb;   // displacement, k_d = 10.05 kN/m^3
var C_di = 1.06 * (draft / d_f) - 0.06;   // C_di = 1.06 (d_i/d_f) - 0.06
var GM = GM_full;   // GM at the draft evaluated, filled in below
if      (draft / d_f >= 1)   GM = GM_full;
else if (draft / d_f >= 0.9) GM = 1.1 * GM_full;
else if (draft / d_f >= 2/3) GM = 1.5 * GM_full;
else if (draft / d_f >= 0.5) GM = 2.0 * GM_full;
else                         GM = GM_full;
var T_r = 2 * k_r * Math.pow(GM / 1 m, -0.5) / 1 m * 1 s;   // T_r, the roll natural period, in seconds
var theta = C_R * beta_RMO * k_esf * (35 - k_theta * C_di * Delta / 1000 kN) * 1 deg;   // theta, the roll amplitude
if      (T_r > 20 s)              theta = theta;
else if (T_r >= 12.5 s)           theta = theta * (1.5375 - 0.027 * T_r / 1 s);
else                              theta = theta * (0.8625 + 0.027 * T_r / 1 s);
if (theta > 30 deg) theta = 30 deg;
else                theta = theta;
var k_v = Math.pow(1 + 0.65 * Math.pow(5.3 - 45 / (L / 1 m), 2) * Math.pow(x / L - 0.45, 2), 0.5);   // k_v (5.7.1(c))
var C_v = Math.cos(mu_deg) + (1 + 2.4 * z / B) * Math.sin(mu_deg) / k_v;   // C_v (5.7.1(c))
var C_l = 0.35 - 0.0005 * (L / 1 m - 200);   // C_l (5.7.1(c))
var k_l = 0.5 + 8 * y / L;   // k_l (5.7.1(c)), the VERTICAL distribution
var C_t = 1.27 * Math.pow(1 + 1.52 * Math.pow(x / L - 0.45, 2), 0.5);   // C_t (5.7.1(c))
var k_t = 0.35 + y / B;   // k_t (5.7.1(c)), the VERTICAL distribution
var a_v = C_v * beta_VAC * k_esf * k_v * a_o / 1 m * g;   // a_v, positive downward
var a_l = C_l * beta_LAC * k_esf * k_l * a_o / 1 m * g;   // a_l, positive forward
var a_t = C_t * beta_TAC * k_esf * k_t * a_o / 1 m * g;   // a_t, positive starboard
var rho_g = rho_tank * g;   // the liquid specific weight (5.7.2(a))
var eta_overflow_add = Math.max(0.6666667 * eta_overflow, 760 mm);   // Fig. 12: 2/3 of the overflow distance, at least 760 mm
var eta = h_tank + eta_deck + (eta_overflow > 0 m ? eta_overflow_add : 0 m) - z;   // eta, measured down from the tank top
var k_s = 1;   // k_s, filled in below
if (tank_is_ballast != 0 || member_11_17 == 0) {
    k_s = 1;
} else {
    k_s = 0.878 + (rho_g - rho_g_lo) / (rho_g_hi - rho_g_lo) * 0.122;
    if (k_s < 0.878) k_s = 0.878;
    else             k_s = k_s;
    if (k_s > 1)     k_s = 1;
    else             k_s = k_s;
}
var p_o = p_vp > 0 N/cm^2 ? Math.max(p_vp - 2.06 N/cm^2, 0 N/cm^2) : 0 N/cm^2;   // p_o = (p_vp - p_n), p_n = 2.06 N/cm^2
var a_i = 0.71 * C_dp * (w_v * a_v + w_l * (l_tank / h_tank) * a_l + w_t * (b_tank / h_tank) * a_t);   // a_i (5.7.2(a))
var theta_e = 0.71 * C_theta * theta;   // theta_e (5.7.2(b))
var phi_e = 0.71 * C_phi * phi;   // phi_e (5.7.2(b))
var zeta = Math.abs(b_tank / 2 - z);   // zeta, from the tank centre line
var zeta_e = zeta;   // zeta_e, filled in below
var eta_e = eta;   // eta_e, filled in below
var dh_i = 0 m;   // delta h_i, filled in below
if (phi_e < 0 deg && theta_e > 0 deg) {
    // branch i of 5.7.2(b)
    zeta_e = b_tank - zeta;
    eta_e = eta;
    dh_i = xi * Math.sin(-phi_e) + C_ru * (zeta_e * Math.sin(theta_e) * Math.cos(phi_e)
         + eta_e * Math.cos(theta_e) * Math.cos(phi_e) - eta);
} else {
    // branch ii of 5.7.2(b)
    zeta_e = zeta - delta_b;
    eta_e = eta - delta_h;
    dh_i = (l_tank - xi) * Math.sin(-phi_e) + C_ru * (zeta_e * Math.sin(theta_e) * Math.cos(phi_e)
         + eta_e * Math.cos(theta_e) * Math.cos(phi_e) - eta);
}
var h_d = k_c * (eta * a_i / g + dh_i);   // h_d = k_c (eta a_i / g + delta h_i) (5.7.2(a))
var p_s = k_s * rho_g * eta;   // static pressure (5.7.2(a))
var p_d = k_s * rho_g * k_u * h_d;   // dynamic pressure (5.7.2(a))
return p_s + p_d + p_o;   // the pressure the function returns
