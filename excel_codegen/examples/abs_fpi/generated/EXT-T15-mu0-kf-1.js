// EXT-T15-mu0-kf-1
// The rules live in the RULE frame: x from the A.P., y from the centre line
// positive starboard, z above the baseline.  The model usually is not, so the
// three mapping cells (ax_* / neg_*) convert it.  All three coordinates go
// through temporaries first: assigning them in place would read a coordinate
// that has already been overwritten.
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
var r_b = 0.703 m;
var freeboard = 0.5 m;
var rho_sea = 1025 kg/m^3;
var g = 9.81 m/s^2;
var Cb = 1;
// --- this case
var draft = 15 m;
var mu_deg = 0 deg;
var k_c = 0.5;
var k_f0 = -1;
var k_u = 1.1;
var beta_EPS = 0.733;
var beta_EPP = 0.733;
var k_esf = 1.1;
var x_o = 51 m;
// --- the calculation
var u = mu_deg;   // the heading, as an Angle -- cos/sin take it directly
var C_1 = 0;   // wave coefficient, 0 until a length bracket matches
// C_1 brackets of 3-2-1/3.5, closed at both ends; outside them C_1 is zero
if      (L >= 61 m  && L <= 90 m ) C_1 = 0.044 * (L / 1 m) + 3.75;
else if (L >  90 m  && L <  300 m) C_1 = 10.75 - Math.pow((300 m - L) / 100 m, 1.5);
else if (L >= 300 m && L <  350 m) C_1 = 10.75;
else if (L >= 350 m && L <= 500 m) C_1 = 10.75 - Math.pow((L - 350 m) / 150 m, 1.5);
else                               C_1 = 0;
var h_do = 1.36 * C_1 * 1 m;   // reference hydrodynamic head, k = 1 (5.5.1)
var z_bilge = r_b;   // rule z of the bilge (the bilge radius)
var z_freeboard = D - freeboard;   // rule z of the highest deck at side
var xL = x / L;   // x / L, dimensionless
var k_lo = 1;   // k_lo = 1.0 amidships
// Fig. 9: k_lo = 1 amidships, rising towards both ends
if      (xL < 0.2) k_lo = 1.5 - xL * 2.5;
else if (xL > 0.7) k_lo = 1 + (xL - 0.7) * 5;
else               k_lo = 1;
var k_l = 1 + (k_lo - 1) * Math.cos(u);   // k_l = 1 + (k_lo - 1) cos mu; GeniE converts the angle inside cos
var k_f = k_f0 * (1 - (1 - Math.cos((x - x_o) / L * 360 deg)) * Math.cos(u));   // k_f, the reference station against the wave
var alpha_1 = (1 - 0.25 * Math.cos(u)) * k_f;   // i = 1, at the waterline, starboard
var alpha_2 = (0.40 - 0.10 * Math.cos(u)) * k_f;   // i = 2, at the bilge, starboard
var alpha_3 = (0.30 - 0.20 * Math.sin(u)) * k_f;   // i = 3, at the bottom centre line
var alpha_5 = (0.75 - 1.25 * Math.sin(u)) * k_f;   // i = 5, at the waterline, port
var alpha_4 = 2 * alpha_3 - alpha_2;   // i = 4 = 2 alpha_3 - alpha_2, at the bilge, port
var h_d1 = k_l * alpha_1 * h_do;   // hydrodynamic head at i = 1
var h_d5 = k_l * alpha_5 * h_do;   // hydrodynamic head at i = 5
var h_star_s = Math.min(z_freeboard - draft, k_u * h_d1);   // Fig. 10: the head above the waterline, capped at the freeboard, starboard
var h_star_p = Math.min(z_freeboard - draft, k_u * h_d5);   // Fig. 10: the head above the waterline, capped at the freeboard, port
var alpha_i = 0;   // alpha at the point evaluated, assigned by the girth branches
// Fig. 8: alpha at the point evaluated
if (z >= z_bilge) {
    if (z <= draft) {
        // bilge to waterline, on the side of the girth the point is on
        if (y >= 0 m) alpha_i = alpha_2 + (alpha_1 - alpha_2) * (z - z_bilge) / (draft - z_bilge);
        else          alpha_i = alpha_4 + (alpha_5 - alpha_4) * (z - z_bilge) / (draft - z_bilge);
    } else {
        // above the waterline, faded out over that side's h_star
        if (y >= 0 m) {
            if (h_star_s <= 0 m) alpha_i = 0;
            else                 alpha_i = alpha_1 - (z - draft) / h_star_s * alpha_1;
        } else {
            if (h_star_p <= 0 m) alpha_i = 0;
            else                 alpha_i = alpha_5 - (z - draft) / h_star_p * alpha_5;
        }
    }
} else {
    // below the bilge: linear across the flat of bottom
    alpha_i = alpha_4 + (alpha_2 - alpha_4) * (z + z_bilge) / (2 * z_bilge);
}
// 5.5.1: which side of the girth, and so which ESF
// the ESF is a property of the SIDE the point is on: EPS to starboard, EPP to port
var ESF_side = 0;
if (y >= 0 m) ESF_side = beta_EPS * k_esf;
else          ESF_side = beta_EPP * k_esf;
var h_di = k_l * alpha_i * h_do;   // hydrodynamic head at the point evaluated (5.5.1)
var h_de = k_c * h_di;   // h_de = k_c h_di (5.5.1 nominal form)
var p_s = Math.max(rho_sea * g * (draft - z), 0 Pa);   // still-water pressure, zero above the waterline
var p_d = rho_sea * g * ESF_side * k_u * h_de;   // hydrodynamic pressure, with the ESF of the side the point is on
return Math.max(p_s + p_d, 0 Pa);   // the pressure the function returns
