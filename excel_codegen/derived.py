"""派生参数（``derived:``）：让一个参数直接引用**同 Case 的其他参数**与**全局参数**。

用途
----
* 中间量：``h_de = k_c * h_di``，模板里直接 ``{{ h_de }}``，不用把这行算式写进模板；
* 需要"看得到数值"的换算结果：参数表里就能读出它，不必等生成完代码再看；
* 链式：派生参数可以引用派生参数（``a -> b -> c``），工具按依赖顺序求值。

限制（**只按需求做这两条**）
--------------------------
* 只能引用 ``global`` 与**同一个 Case** 的 ``local``；不跨 Case、不跨工作簿；
* ``global`` 的派生参数不能引用 ``local``（那时还没有"当前 Case"这个概念）。

两条路径，一套语义
------------------
* **Python 侧**（快照渲染 / 导出文件 / 指纹）：用 ``jinja_env`` 的环境求值，拿真正的值；
* **Excel 侧**（参数表里那一格）：能翻译成 Excel 公式的就写公式（改输入自动重算），
  翻译不了（用了 ``|`` 过滤器里没有对应的、条件表达式以外的花活……）就把算好的值写进去，
  并告警"改输入后需要重跑 --write-excel"。

翻译覆盖的子集：``+ - * / **``、``%``（-> ``MOD``）、``~``（-> ``&``，字符串拼接）、
括号、数字/字符串/布尔常量、变量、``min/max/abs/round/int/float``、
``|round|abs|int|float``、比较运算、``and/or/not``、条件表达式（-> ``IF``）。
其余一律报错（不猜），由调用方降级。
"""

from __future__ import annotations

from typing import Any, Iterable, Mapping, MutableMapping, Sequence

from jinja2 import Environment, TemplateError, TemplateSyntaxError, UndefinedError, meta, nodes

from .formula import lookup_expr
from .jinja_env import build_environment
from .models import ProjectConfig, VariableDef
from .utils import RenderError, to_text

__all__ = [
    "DerivedError",
    "DerivedNotTranslatable",
    "derived_variables",
    "evaluate_derived",
    "expression_names",
    "is_derived",
    "is_translatable",
    "ordered_derived",
    "to_excel",
    "untranslatable_names",
    "validate_config",
]


class DerivedError(RenderError):
    """派生参数写得不对：语法错、引用越界、循环引用、类型不对……"""


class DerivedNotTranslatable(DerivedError):
    """表达式本身没错，但超出"能翻译成 Excel 公式"的子集。

    与 :class:`DerivedError` 分开，是为了让调用方**只**对这种情况降级
    （往格子里写算好的值），而不是把语法错误也静默吞掉。
    """


def is_derived(variable: VariableDef) -> bool:
    return bool(variable.derived)


def derived_variables(variables: Iterable[VariableDef]) -> list[VariableDef]:
    return [variable for variable in variables if is_derived(variable)]


# --------------------------------------------------------------------------- #
# 表达式：名字与顺序
# --------------------------------------------------------------------------- #
def expression_names(expression: str, *, env: Environment | None = None) -> set[str]:
    """表达式里引用到的变量名。"""
    environment = env or build_environment()
    try:
        ast = environment.parse("{{ " + expression + " }}")
    except TemplateSyntaxError as exc:
        raise DerivedError(
            f"派生表达式语法错误（第 {exc.lineno} 行）：{exc.message}；表达式：{expression!r}"
        ) from exc
    return set(meta.find_undeclared_variables(ast))


def ordered_derived(
    variables: Sequence[VariableDef],
    *,
    scope: str,
    env: Environment | None = None,
) -> list[VariableDef]:
    """按依赖顺序排列派生参数；有环就报错（带上环上的名字）。"""
    environment = env or build_environment()
    derived = {variable.name: variable for variable in derived_variables(variables)}
    if not derived:
        return []

    deps = {
        name: expression_names(variable.derived or "", env=environment) & derived.keys()
        for name, variable in derived.items()
    }

    ordered: list[VariableDef] = []
    done: set[str] = set()
    while len(ordered) < len(derived):
        ready = [name for name in derived if name not in done and deps[name] <= done]
        if not ready:
            cycle = _find_cycle(deps)
            raise DerivedError(
                f"{scope} 的派生参数存在循环引用：{' -> '.join(cycle)}；"
                "派生参数只能引用同 Case 的其他参数与全局参数，不能互相依赖成环"
            )
        for name in sorted(ready):
            ordered.append(derived[name])
            done.add(name)
    return ordered


def _find_cycle(deps: Mapping[str, set[str]]) -> list[str]:
    """从依赖图里挑出一个环，用于报错信息。"""
    visiting: list[str] = []
    seen: set[str] = set()

    def walk(node: str) -> list[str] | None:
        if node in visiting:
            return visiting[visiting.index(node):] + [node]
        if node in seen:
            return None
        visiting.append(node)
        for dep in sorted(deps.get(node, ())):
            found = walk(dep)
            if found:
                return found
        visiting.pop()
        seen.add(node)
        return None

    for name in sorted(deps):
        found = walk(name)
        if found:
            return found
    return ["<未知>"]


# --------------------------------------------------------------------------- #
# 求值（Python 侧）
# --------------------------------------------------------------------------- #
def evaluate_derived(
    variables: Sequence[VariableDef],
    context: MutableMapping[str, Any],
    *,
    scope: str,
    env: Environment | None = None,
    forbidden: Mapping[str, str] | None = None,
) -> None:
    """按依赖顺序求值派生参数，结果写进 ``context``（原地更新）。

    :param scope: 出错信息里用的作用域描述，例如 ``"global"`` / ``"Case1 的 local"``
    :param forbidden: ``{变量名: 说明}``；表达式引用到这些名字就报错
        （用来实现"global 派生参数不能引用 local 变量"）
    """
    environment = env or build_environment()
    forbidden = forbidden or {}
    for variable in ordered_derived(variables, scope=scope, env=environment):
        expression = variable.derived or ""
        for name in sorted(expression_names(expression, env=environment)):
            if name in forbidden:
                raise DerivedError(
                    f"派生参数 {variable.name!r}（{scope}）引用了 {name!r}：{forbidden[name]}；"
                    f"表达式：{expression!r}"
                )
        context[variable.name] = _evaluate(expression, context, env=environment, variable=variable)


def _evaluate(
    expression: str,
    context: Mapping[str, Any],
    *,
    env: Environment,
    variable: VariableDef,
) -> Any:
    try:
        compiled = env.compile_expression(expression, undefined_to_none=False)
    except TemplateSyntaxError as exc:
        raise DerivedError(
            f"派生参数 {variable.name!r} 的表达式语法错误（第 {exc.lineno} 行）：{exc.message}；"
            f"表达式：{expression!r}"
        ) from exc
    try:
        return compiled(**context)
    except UndefinedError as exc:
        raise DerivedError(
            f"派生参数 {variable.name!r} 引用了取不到的变量：{exc.message or exc}；"
            f"表达式：{expression!r}"
            "（派生参数只能引用 YAML 里定义的 global / 同 Case 的 local）"
        ) from exc
    except TemplateError as exc:
        raise DerivedError(f"派生参数 {variable.name!r} 求值失败：{exc}；表达式：{expression!r}") from exc
    except Exception as exc:  # noqa: BLE001 - 表达式里的任意 Python 异常
        raise DerivedError(
            f"派生参数 {variable.name!r} 求值失败：{type(exc).__name__}: {exc}；表达式：{expression!r}"
        ) from exc


# --------------------------------------------------------------------------- #
# 翻译成 Excel 公式
# --------------------------------------------------------------------------- #
_BINOPS = {"+": "+", "-": "-", "*": "*", "/": "/", "**": "^"}
#: Jinja 的比较运算用的是名字（gt / lteq …），不是符号
_COMPARE = {
    "eq": "=", "ne": "<>", "gt": ">", "lt": "<", "gteq": ">=", "lteq": "<=",
    "==": "=", "!=": "<>", ">=": ">=", "<=": "<=", ">": ">", "<": "<",
}
#: 表达式里的函数 -> Excel 函数。``int`` 用 ``TRUNC``：Python 的 ``int()`` 向零截断，
#: 而 Excel 的 ``INT()`` 是向下取整（``int(-2.5)`` 在两边会差 1）。
_FUNCS = {"min": "MIN", "max": "MAX", "abs": "ABS", "int": "TRUNC", "float": None}
_FILTERS = {"abs": "ABS", "int": "TRUNC", "float": None, "string": None}

#: 两种语言语义不同的运算：**故意不翻译**（翻译了就会"Excel 里看到的"与"导出的"不一致）
_NOT_TRANSLATABLE = {
    "round": "Python 的 round() 是银行家舍入（round(2.5) == 2），Excel 的 ROUND() 是四舍五入（2.5 -> 3）",
    "ceil": "Python 的 ceil() 与 Excel 的 CEILING 边界行为不同",
    "floor": "Python 的 floor() 与 Excel 的 FLOOR 边界行为不同（负数）",
}


def to_excel(
    expression: str,
    *,
    name: str,
    resolve,
    env: Environment | None = None,
    condition: bool = False,
) -> str:
    """把派生表达式翻译成 Excel 公式（不带前导 ``=``）。

    :param resolve: ``变量名 -> Excel 引用`` 的回调（由调用方决定列/表名与相对/绝对）
    :param condition: 这是不是一个**条件**表达式。为 ``True`` 时 ``and`` / ``or`` / ``not``
        可以翻译成 ``AND`` / ``OR`` / ``NOT``（Python 的 ``and``/``or`` 有返回值语义，
        与 Excel 不同，所以只在条件位置放行）。公式模式的行内 ``{% if %}`` 走这条路。
    :raises DerivedError: 表达式超出可翻译子集（调用方应当降级为"写入算好的值"）
    """
    environment = env or build_environment()
    try:
        ast = environment.parse("{{ " + expression + " }}")
    except TemplateSyntaxError as exc:
        raise DerivedError(
            f"派生参数 {name!r} 的表达式语法错误：{exc.message}；表达式：{expression!r}"
        ) from exc
    body = getattr(ast, "body", [])
    if len(body) != 1 or not isinstance(body[0], nodes.Output) or len(body[0].nodes) != 1:
        raise DerivedError(f"派生参数 {name!r} 的表达式不是单个表达式：{expression!r}")
    return _translate(
        body[0].nodes[0], name=name, resolve=resolve, env=environment, condition=condition
    )


def _translate(node, *, name: str, resolve, env: Environment, condition: bool) -> str:
    if isinstance(node, nodes.Const):
        return _excel_literal(node.value)
    if isinstance(node, nodes.Name):
        return resolve(node.name)
    if isinstance(node, nodes.Getattr):
        # 属性访问（``x.value`` / ``x.prefix`` …）：把属性名一并交给调用方 —— 只有它知道
        # 该怎么把一个"取值 / 前缀"映射成单元格引用。老的回调只收一个参数，用 TypeError 兜住。
        inner = node.node
        if not isinstance(inner, nodes.Name):
            raise DerivedNotTranslatable(
                f"表达式 {name!r} 里的 .{node.attr} 只能接在变量名后面"
            )
        try:
            return resolve(inner.name, node.attr)
        except TypeError as exc:
            raise DerivedNotTranslatable(
                f"表达式 {name!r} 里用了属性访问 .{node.attr}，但当前上下文不支持"
                "（属性访问只在公式模式的行内 {% if %} 条件里可用）"
            ) from exc
    if isinstance(node, nodes.Concat):
        return "&".join(
            f"({_translate(item, name=name, resolve=resolve, env=env, condition=False)})"
            for item in node.nodes
        )
    # ⚠ and / or 必须排在 BinExpr **前面**：jinja2 里 ``And`` / ``Or`` 是 ``BinExpr`` 的子类，
    # 先命中 BinExpr 分支的话它们会被当成"不支持的运算符"（这段曾经是死代码，见 0.6.0 修复）。
    if isinstance(node, (nodes.And, nodes.Or)):
        if not condition:
            raise DerivedNotTranslatable(
                f"派生参数 {name!r} 的 and/or 只能用在条件表达式里"
                "（Python 的 and/or 返回值语义与 Excel 不同，避免静默偏差）"
            )
        function = "AND" if isinstance(node, nodes.And) else "OR"
        left = _translate(node.left, name=name, resolve=resolve, env=env, condition=True)
        right = _translate(node.right, name=name, resolve=resolve, env=env, condition=True)
        return f"{function}({left},{right})"
    if isinstance(node, nodes.BinExpr):
        operator = getattr(node, "operator", None)
        left = _translate(node.left, name=name, resolve=resolve, env=env, condition=False)
        right = _translate(node.right, name=name, resolve=resolve, env=env, condition=False)
        if operator == "%":
            return f"MOD({left},{right})"
        if operator == "//":
            raise DerivedNotTranslatable(
                f"派生参数 {name!r} 用了 //（整除）：Excel 没有等价运算，请改用 / 或 round()"
            )
        if operator not in _BINOPS:
            raise DerivedNotTranslatable(f"派生参数 {name!r} 用了不支持的运算符 {operator!r}")
        return f"({left}{_BINOPS[operator]}{right})"
    if isinstance(node, nodes.Neg):
        return f"-({_translate(node.node, name=name, resolve=resolve, env=env, condition=False)})"
    if isinstance(node, nodes.Pos):
        return f"({_translate(node.node, name=name, resolve=resolve, env=env, condition=False)})"
    if isinstance(node, nodes.Not):
        if not condition:
            raise DerivedNotTranslatable(f"派生参数 {name!r} 的 not 只能用在条件表达式里")
        inner = _translate(node.node, name=name, resolve=resolve, env=env, condition=True)
        return f"NOT({inner})"
    if isinstance(node, nodes.Compare):
        return _translate_compare(node, name=name, resolve=resolve, env=env)
    if isinstance(node, nodes.CondExpr):
        test = _translate(node.test, name=name, resolve=resolve, env=env, condition=True)
        yes = _translate(node.expr1, name=name, resolve=resolve, env=env, condition=False)
        no = _translate(node.expr2, name=name, resolve=resolve, env=env, condition=False)
        return f"IF({test},{yes},{no})"
    if isinstance(node, nodes.Filter):
        return _translate_filter(node, name=name, resolve=resolve, env=env)
    if isinstance(node, nodes.Call):
        return _translate_call(node, name=name, resolve=resolve, env=env)
    raise DerivedNotTranslatable(
        f"派生参数 {name!r} 的表达式超出可翻译范围（{type(node).__name__}）："
        "能翻译的子集见 docs/template_guide.md §15；"
        "也可以让它留在 Python 侧（表里写入算好的值，改输入后重跑 --write-excel）"
    )


def _translate_compare(node, *, name: str, resolve, env: Environment) -> str:
    parts: list[str] = []
    left = _translate(node.expr, name=name, resolve=resolve, env=env, condition=False)
    for operand in node.ops:
        operator = getattr(operand, "op", None)
        right = _translate(operand.expr, name=name, resolve=resolve, env=env, condition=False)
        if operator not in _COMPARE:
            raise DerivedNotTranslatable(f"派生参数 {name!r} 用了不支持的比较运算 {operator!r}")
        parts.append(f"({left}{_COMPARE[operator]}{right})")
        left = right
    return parts[0] if len(parts) == 1 else "AND(" + ",".join(parts) + ")"


def _translate_filter(node, *, name: str, resolve, env: Environment) -> str:
    if node.dyn_args is not None or node.dyn_kwargs is not None:
        raise DerivedNotTranslatable(f"派生参数 {name!r} 的过滤器用了动态参数，无法翻译")
    if node.name in _NOT_TRANSLATABLE:
        raise DerivedNotTranslatable(
            f"派生参数 {name!r} 用了 |{node.name}：{_NOT_TRANSLATABLE[node.name]}，"
            "所以不往 Excel 公式翻译（表里会写入 Python 算好的值）"
        )
    function = _FILTERS.get(node.name)
    if function is None and node.name not in _FILTERS:
        raise DerivedNotTranslatable(f"派生参数 {name!r} 用了过滤器 |{node.name}，Excel 侧没有对应函数")
    value = _translate(node.node, name=name, resolve=resolve, env=env, condition=False)
    if function is None:  # float / string：Excel 里就是原值
        return f"({value})"
    return f"{function}({value})"


def _translate_call(node, *, name: str, resolve, env: Environment) -> str:
    if not isinstance(node.node, nodes.Name):
        raise DerivedNotTranslatable(f"派生参数 {name!r} 调用了非函数对象，无法翻译")
    if node.kwargs and any(kw.key for kw in node.kwargs):
        raise DerivedNotTranslatable(f"派生参数 {name!r} 的函数用了关键字参数，无法翻译")
    if node.dyn_args is not None or node.dyn_kwargs is not None:
        raise DerivedNotTranslatable(f"派生参数 {name!r} 的函数用了 *args/**kwargs，无法翻译")

    function_name = node.node.name
    if function_name in _NOT_TRANSLATABLE:
        raise DerivedNotTranslatable(
            f"派生参数 {name!r} 用了 {function_name}()：{_NOT_TRANSLATABLE[function_name]}，"
            "所以不往 Excel 公式翻译（表里会写入 Python 算好的值）"
        )
    if function_name not in _FUNCS:
        raise DerivedNotTranslatable(
            f"派生参数 {name!r} 调用了 {function_name}()，Excel 侧没有对应函数"
            "（支持 min/max/abs/int/float）"
        )
    arguments = [
        _translate(argument, name=name, resolve=resolve, env=env, condition=False)
        for argument in node.args
    ]
    excel_function = _FUNCS[function_name]
    if excel_function is None:  # float()：Excel 里数字就是数字
        if len(arguments) != 1:
            raise DerivedNotTranslatable(f"派生参数 {name!r} 的 {function_name}() 只支持一个参数")
        return f"({arguments[0]})"
    if not arguments:
        raise DerivedNotTranslatable(f"派生参数 {name!r} 的 {function_name}() 缺参数")
    return f"{excel_function}({','.join(arguments)})"


def _excel_literal(value: Any) -> str:
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return repr(value)
    text = to_text(value)
    return '"' + text.replace('"', '""') + '"'


def is_translatable(variable: VariableDef, *, env: Environment | None = None) -> bool:
    """这个派生表达式能不能写成 Excel 公式（不能就只能往格子里写算好的值）。"""
    if not is_derived(variable):
        return True
    try:
        to_excel(variable.derived or "", name=variable.name, resolve=lambda _: "A1", env=env)
        return True
    except DerivedNotTranslatable:
        return False


def untranslatable_names(config: ProjectConfig, *, env: Environment | None = None) -> list[str]:
    """哪些派生参数翻译不成 Excel 公式（表里只能写入算好的值）。"""
    environment = env or build_environment()
    return [
        variable.name
        for variable in [*config.global_variables, *config.local_variables]
        if is_derived(variable) and not is_translatable(variable, env=environment)
    ]


def validate_config(config: ProjectConfig, *, env: Environment | None = None) -> list[str]:
    """配置期的派生参数检查：语法、引用范围、循环；返回警告（不能翻译成 Excel 公式的）。

    :raises DerivedError: 表达式语法错 / 引用了不该引用的名字 / 循环引用 / 引用了未定义的名字
    """
    environment = env or build_environment()
    global_names = set(config.global_names)
    local_names = set(config.local_names)
    known = global_names | local_names

    for variable in derived_variables(config.global_variables):
        for name in sorted(expression_names(variable.derived or "", env=environment)):
            if name in local_names:
                raise DerivedError(
                    f"派生参数 {variable.name!r}（global）引用了 local 变量 {name!r}："
                    "global 的派生参数不能引用 local 变量（那时还没有当前 Case）"
                )
            if name not in known:
                raise DerivedError(
                    f"派生参数 {variable.name!r}（global）引用了未定义的变量 {name!r}；"
                    "派生参数只能引用 YAML 里定义的 global / 同 Case 的 local"
                )

    for variable in derived_variables(config.local_variables):
        for name in sorted(expression_names(variable.derived or "", env=environment)):
            if name not in known:
                raise DerivedError(
                    f"派生参数 {variable.name!r}（local）引用了未定义的变量 {name!r}；"
                    "派生参数只能引用 YAML 里定义的 global / 同 Case 的 local"
                )

    ordered_derived(config.global_variables, scope="global 的派生参数", env=environment)
    ordered_derived(config.local_variables, scope="local 的派生参数", env=environment)

    untranslatable = untranslatable_names(config, env=environment)
    return (
        [
            "这些派生参数写不成 Excel 公式，参数表里只能写入算好的值"
            f"（改输入后要重跑 --write-excel 刷新）：{', '.join(untranslatable)}"
        ]
        if untranslatable
        else []
    )
