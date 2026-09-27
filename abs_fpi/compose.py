"""把若干"规则集 YAML"合成一本项目工作簿的 YAML。

有点历史了：这个脚本写在工具还没有 ``extends`` 的时候（见指南 §16）。现在大多数场景
直接在项目 YAML 里写 ``extends: [a.yaml, b.yaml]`` 就行 —— 合并判据与此处一致，
而且 ``template_file`` 会相对各自文件解析。

它还留着的原因是**它比 extends 多做一件事**：按规则集后缀改写 ``output_sheet``
（``Code`` → ``Code EXT``），适合"规则集不关心自己叫哪张表"的批量场景。
另外它是工具变更的**验收证据**之一（`build.py` 会跑它）。

为什么要它
----------
工具只认**一个 YAML = 一本工作簿**。而一个项目通常要在一本工作簿里放几套规范，
共用一张 Global 表（船的主尺度只填一次）。没有合并能力，只能把几套变量**手抄**进一个文件
—— 那是漂移的温床。

于是这里用"合成"代替"手抄"：

    abs_fpi_external.yaml ─┐
                           ├─ compose.py ──► abs_fpi.yaml ── init/render ──► 一本工作簿
    abs_fpi_internal.yaml ─┘

每个规则集仍是**唯一真源**（变量 + 模板），项目 YAML 是生成物。合成时会检查：

* 同名变量的 ``prefix`` / ``suffix`` / ``type`` 必须一致 —— 不一致直接报错
  （这三个决定生成出来的文本，不能悄悄让某一个赢）；
* ``default`` / ``description`` 不一致只**告警**（不同规则集的示例工况本来就不同）；
* 模板名重复必须内容一致，否则报错。

用法::

    python compose.py            # 读 compose.yaml → 写 abs_fpi.yaml
"""

from __future__ import annotations

import sys
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from excel_codegen import load_config  # noqa: E402

MANIFEST = HERE / "compose.yaml"

#: 这三个字段决定生成出来的文本，同名变量之间必须逐字一致
STRICT_FIELDS = ("prefix", "suffix", "type")
#: 这些字段不一致只告警
LOOSE_FIELDS = ("default", "description")


def read_manifest(path: Path) -> dict:
    with path.open(encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def merge_variables(target: dict, incoming: list[dict], *, scope: str, source: str, warnings: list[str]) -> None:
    by_name = {item["name"]: item for item in target[scope]}
    for item in incoming:
        name = item["name"]
        other = by_name.get(name)
        if other is None:
            target[scope].append(dict(item))
            by_name[name] = target[scope][-1]
            continue
        for field in STRICT_FIELDS:
            a, b = other.get(field, ""), item.get(field, "")
            if a != b:
                raise SystemExit(
                    f"变量 {name!r} 在 {scope} 里被两个规则集定义成不同的 {field}："
                    f"{a!r}（{other.get('_source')}）vs {b!r}（{source}）。"
                    "这三个字段决定生成出来的文本，必须先统一。"
                )
        for field in LOOSE_FIELDS:
            if other.get(field) != item.get(field):
                warnings.append(
                    f"变量 {name!r}（{scope}）的 {field} 不一致：保留第一个"
                    f"（{other.get(field)!r}），忽略 {source} 的 {item.get(field)!r}"
                )


def compose(manifest_path: Path) -> Path:
    manifest = read_manifest(manifest_path)
    base = manifest_path.parent

    merged: dict = {"version": 1, "variables": {"global": [], "local": []}, "templates": []}
    outputs: list[str] = []
    warnings: list[str] = []

    for entry in manifest["rulesets"]:
        rel = entry["file"]
        cfg_path = (base / rel).resolve()
        # 先让工具自己校验一遍：规则集 YAML 单独用也必须是对的
        load_config(cfg_path)
        with cfg_path.open(encoding="utf-8") as handle:
            raw = yaml.safe_load(handle)
        tag = entry.get("set") or cfg_path.stem
        filt = entry.get("filter")

        for scope in ("global", "local"):
            items = (raw.get("variables") or {}).get(scope) or []
            for item in items:
                item["_source"] = rel
            merge_variables(merged["variables"], items, scope=scope, source=rel, warnings=warnings)

        for tpl in raw.get("templates") or []:
            tpl = dict(tpl)
            # 两套规则的输出表不能同名，按规则集后缀区分
            sheet = tpl.get("output_sheet", "Output")
            tpl["output_sheet"] = entry.get("sheet_prefix", "") + sheet + entry.get("sheet_suffix", f" {tag}")
            if filt and len(manifest["rulesets"]) > 1:
                tpl["case_filter"] = filt
            same = [t for t in merged["templates"] if t["name"] == tpl["name"]]
            if same:
                if yaml.safe_dump(same[0], sort_keys=True) != yaml.safe_dump(tpl, sort_keys=True):
                    raise SystemExit(f"模板名 {tpl['name']!r} 在两个规则集里内容不同")
                continue
            merged["templates"].append(tpl)
            if tpl["output_sheet"] not in outputs:
                outputs.append(tpl["output_sheet"])

    for scope in ("global", "local"):
        for item in merged["variables"][scope]:
            item.pop("_source", None)

    project = {
        "version": 1,
        "excel": {
            "output": manifest["output"],
            "template_sheet": "Template",
            "howto_sheet": "HOWTO",
            "sheets": {
                "global": "Global Parameter",
                "local": "Local Parameter",
                "outputs": outputs,
            },
        },
        "variables": merged["variables"],
        "templates": merged["templates"],
    }

    out_path = base / manifest.get("project", "abs_fpi.yaml")
    header = (
        "# =========================================================================== #\n"
        "# 本文件由 compose.py 生成，请不要手改 —— 改了会在下次 compose 时被覆盖。\n"
        "#\n"
        "# 真源是各规则集 YAML（见 compose.yaml 的 rulesets）：\n"
        + "".join(f"#   - {e['file']}\n" for e in manifest["rulesets"])
        + "# 改参数请改工作簿；加规则集请加规则集 YAML 再跑 python compose.py。\n"
        "# =========================================================================== #\n"
    )
    with out_path.open("w", encoding="utf-8") as handle:
        handle.write(header)
        yaml.safe_dump(project, handle, sort_keys=False, allow_unicode=True, width=100)

    # 合成结果自己也要过一遍工具的校验
    load_config(out_path)

    for w in warnings:
        print(f"  ! {w}")
    print(
        f"wrote {out_path.name}: {len(project['variables']['global'])} 全局 / "
        f"{len(project['variables']['local'])} 局部变量, {len(project['templates'])} 模板 "
        f"→ {', '.join(outputs)}"
    )
    return out_path


def main() -> int:
    compose(MANIFEST)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
