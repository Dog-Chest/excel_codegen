/* ============================================================================
 * compare_with_rules.js
 * ----------------------------------------------------------------------------
 * 交叉校验：把 spreadsheet_codegen 渲染出来的 GeniE 代码，与 GeniE/Rules 那套已经过
 * 量纲/结构校验的生成器对同一批工况产出的代码逐行比对。
 *
 * 为什么需要它：spreadsheet_codegen 只做「文本替换」，它不认识 GeniE 的量纲、
 * 不知道 `if` 必须配 `else`、也不会检查 `sin/cos` 收到的是不是角度。
 * 所以移植之后必须用另一套有检查的工具来确认移植没有走样 —— 这个脚本就是那条链。
 *
 *   node compare_with_rules.js                   # 缺外部仓库 → SKIP，退出码 2
 *   node compare_with_rules.js --allow-missing   # 缺外部仓库 → SKIP，退出码 0
 *
 * 外部依赖：同级目录的 `../../GeniE/Rules`（genie.js / build.js）。它是独立仓库，
 * 不在本工作区里，所以缺少时必须**明确跳过**而不是静默通过 —— 否则 CI 会把
 * "没验证过"当成"验证过了"。
 * ==========================================================================*/

'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');

const RULES_DIR = path.resolve(__dirname, '..', '..', 'GeniE', 'Rules');
const RULES_MODULES = ['genie.js', 'build.js'];
const missing = RULES_MODULES.filter(function (name) {
  return !fs.existsSync(path.join(RULES_DIR, name));
});

if (missing.length) {
  const allowMissing = process.argv.indexOf('--allow-missing') !== -1;
  console.log('SKIP  没有做交叉校验：找不到外部依赖 GeniE/Rules');
  console.log('      期望位置：' + RULES_DIR);
  console.log('      缺少文件：' + missing.join(', '));
  console.log('      说明：这是一条**独立**证据（与另一套已校验的生成器逐行比对）；');
  console.log('            该仓库不在本工作区时无法执行。其余验证（spreadsheet-codegen check、');
  console.log('            verify_excel_engine.py、build.py --check）不依赖它。');
  console.log('      退出码：' + (allowMissing ? '0（--allow-missing，按跳过处理）'
                                             : '2（明确跳过，不等于通过）'));
  process.exit(allowMissing ? 0 : 2);
}

const genie = require(path.join(RULES_DIR, 'genie.js'));
const build = require(path.join(RULES_DIR, 'build.js'));

let checks = 0;
let failures = 0;
function check(cond, msg, detail) {
  checks += 1;
  if (cond) return true;
  failures += 1;
  console.log('  FAIL  ' + msg + (detail ? '\n        ' + detail : ''));
  return false;
}

/** 去掉行尾注释、行首注释、空行，再 trim —— 只留"算出来的东西" */
function normalize(lines) {
  const out = [];
  lines.forEach(function (raw) {
    let s = String(raw).replace(/\/\/.*$/, '').trim();
    if (s === '') return;
    // 坐标变换语句两边都排除：
    //   GeniE/Rules 在恒等映射下**不输出**它，而 spreadsheet_codegen 侧为了能进公式模式
    //   （公式模式没有 {% if %}）总是输出。变换本身由 6 个映射格纯替换而来，
    //   正确性由 verify_excel_engine.py --change 与 rule_selfcheck.js 的 48 种映射各自覆盖。
    if (/^(var t[123] = [xyz];|[xyz] = -?t[123];)$/.test(s)) return;
    out.push(s.replace(/\s+/g, ' '));
  });
  return out;
}

function diff(a, b) {
  const n = Math.max(a.length, b.length);
  const bad = [];
  for (let i = 0; i < n; i++) {
    if (a[i] !== b[i]) {
      bad.push('line ' + (i + 1) + ':\n          rules: ' + (a[i] === undefined ? '(none)' : a[i]) +
        '\n          port : ' + (b[i] === undefined ? '(none)' : b[i]));
    }
  }
  return bad;
}

const tmp = path.join(os.tmpdir(), 'grt_port_cmp.xlsx');
const built = build.build({ writeTo: tmp });

const SETS = [
  { set: 'EXT', prefix: 'EXT-', dir: path.join(__dirname, 'generated') },
  { set: 'INT', prefix: '', dir: path.join(__dirname, 'generated') },
];

SETS.forEach(function (s) {
  const block = built.result.blocks.filter(function (b) { return b.set === s.set; })[0];
  console.log('\n== ' + s.set + '  (' + block.labels.length + ' cases)');
  block.labels.forEach(function (label, i) {
    const file = path.join(s.dir, label + '.js');
    if (!fs.existsSync(file)) {
      check(false, s.set + ' ' + label + ': the ported file exists (' + file + ')');
      return;
    }
    const ported = fs.readFileSync(file, 'utf8').split(/\r?\n/);
    const mine = build.wrapFunction(block.code[i], label).split(/\r?\n/);
    const a = normalize(mine);
    const b = normalize(ported);
    const bad = diff(a, b);
    check(bad.length === 0, s.set + ' ' + label + ': identical to the validated rules ' +
      '(' + a.length + ' vs ' + b.length + ' significant lines)',
      bad.slice(0, 4).join('\n        '));

    // 借 GeniE/Rules 的检查跑一遍移植产物
    const ifBad = genie.checkIfElse(ported);
    check(ifBad.length === 0, s.set + ' ' + label + ': every if has an else', ifBad.join(' | '));
    const nameBad = genie.checkNames(ported, { x: true, y: true, z: true });
    check(nameBad.length === 0, s.set + ' ' + label + ': declared before read, never twice',
      nameBad.join(' | '));
  });
});

try { fs.unlinkSync(tmp); } catch (e) { /* gone */ }

console.log('');
console.log(checks + ' checks, ' + failures + ' failure(s)');
process.exit(failures ? 1 : 0);
