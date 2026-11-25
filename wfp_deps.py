# -*- coding: utf-8 -*-
"""依赖自检：按 requirements.txt 全量校验每个发行包是否已安装。

供启动脚本与 Web 入口在启动前调用。缺少依赖时明确报错退出，
而不是让功能静默降级（例如 jieba 缺失会让断行合并悄悄失效）。
"""
from __future__ import annotations

import sys
from importlib import metadata
from pathlib import Path

try:  # packaging 随 pip/setuptools 提供，缺失时退化为手工解析
    from packaging.requirements import InvalidRequirement, Requirement
except ImportError:  # pragma: no cover - 仅在极简环境触发
    InvalidRequirement = Exception
    Requirement = None

_SPECIFIER_SEPARATORS = (">=", "<=", "==", "!=", "~=", ">", "<", "=")


def iter_requirement_lines(path):
    """逐行产出 requirements 中的有效条目（去掉注释、空行与 pip 选项）。"""
    text = Path(path).read_text(encoding="utf-8-sig")
    for raw in text.splitlines():
        line = raw.split("#", 1)[0].strip()
        if not line or line.startswith("-"):
            continue
        yield line


def requirement_name(line):
    """返回该条目的发行包名；环境标记不满足（如 Linux 上的 pywin32）时返回 None。"""
    if Requirement is not None:
        try:
            requirement = Requirement(line)
        except InvalidRequirement:
            return None
        if requirement.marker is not None and not requirement.marker.evaluate():
            return None
        return requirement.name
    head = line.split(";", 1)[0].strip()
    name = head.split("[", 1)[0]
    for separator in _SPECIFIER_SEPARATORS:
        name = name.split(separator, 1)[0]
    return name.strip() or None


def missing_requirements(path):
    """返回 requirements.txt 中尚未安装的发行包名列表（去重、保持原顺序）。"""
    missing = []
    seen = set()
    for line in iter_requirement_lines(path):
        name = requirement_name(line)
        if not name:
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        try:
            metadata.version(name)
        except metadata.PackageNotFoundError:
            missing.append(name)
    return missing


def check_requirements(paths, log=print):
    """全量校验多个依赖清单，返回缺失项列表 [(清单路径, 包名)]。"""
    missing = []
    for path in paths:
        path = Path(path)
        if not path.is_file():
            log(f"[deps] 跳过不存在的依赖清单：{path}")
            continue
        for name in missing_requirements(path):
            missing.append((str(path), name))
    return missing


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    missing = check_requirements(argv or ["requirements.txt"])
    if not missing:
        print("[deps] 依赖检查通过。")
        return 0
    print("[deps] 以下依赖尚未安装：")
    for path, name in missing:
        print(f"  - {name}   （来自 {path}）")
    print("[deps] 请先执行：python -m pip install -r <依赖清单>")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
