# -*- coding: utf-8 -*-
"""技能包质量门禁（通用引擎）——由 skill-cicd 生成，勿手工改动。

单一来源: skills/skill-cicd/scripts/qa_engine.py
升级方式: 改引擎后对已接入仓运行
    python skills/skill-cicd/scripts/gen_cicd.py --dir <仓> --apply
本仓差异只允许出现在 .cicd/config.json。

检查项（全部软探测，包有什么查什么，不强制结构统一）：
  1. .cicd/config.json 存在且含 name
  2. skill.json 存在则须为合法 JSON 且含 version
  3. SKILL.md 存在则 frontmatter(name/description) 完整；若头部含 `> 版本: vX.Y.Z` 且 skill.json 存在则须一致
  4. CHANGELOG（versions/CHANGELOG.md 或 CHANGELOG.md，探测存在）若存在且 skill.json 存在则最新版本行须一致
  5. README.md 若存在且 skill.json 存在则标题版本须一致（正则按 config.name）
  6. scripts/*.py 语法编译（py_compile 全部，含 qa_checks.py 自身）
  7. config.required_files 列出的资源须存在
  8. 敏感信息扫描（ghp_ token / GITHUB_TOKEN= / password= / Authorization）——文本文件；--staged 模式仅扫描暂存文本变更
用法：
  python scripts/qa_checks.py            # 检查工作区
  python scripts/qa_checks.py --staged   # 仅检查 git 暂存区文本变更（同步前门禁）
退出码：0=通过，1=失败
"""
import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / ".cicd" / "config.json"
FAILED = []


def check(name, ok, detail=""):
    tag = "PASS" if ok else "FAIL"
    print(f"[{tag}] {name}" + (f" — {detail}" if detail else ""))
    if not ok:
        FAILED.append(name)


def load_json(p):
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:
        check(f"JSON 合法 {p.relative_to(ROOT)}", False, str(e))
        return None


def text_of(p):
    try:
        return p.read_text(encoding="utf-8")
    except Exception:
        return None


def main():
    cfg = load_json(CONFIG_PATH)
    name = (cfg or {}).get("name") or ROOT.name
    if cfg is None:
        check(".cicd/config.json 存在且含 name", False, "缺失：先运行 skill-cicd/gen_cicd.py 生成")
    else:
        check(".cicd/config.json 存在且含 name", bool(cfg.get("name")), f"name={name}")

    # skill.json 版本基
    sv = load_json(ROOT / "skill.json")
    ver = (sv or {}).get("version") if isinstance(sv, dict) else None
    want = ("v" + ver) if isinstance(ver, str) and ver else None
    if sv is not None:
        check("skill.json 合法且含 version", isinstance(ver, str) and bool(ver), f"version={ver}")

    # SKILL.md
    sk_path = ROOT / "SKILL.md"
    sk = text_of(sk_path) if sk_path.exists() else None
    if sk is None:
        check("SKILL.md 存在", False, "缺失")
    else:
        fm = re.match(r"^---\n(.*?)\n---", sk, re.S)
        fm_ok = bool(fm and re.search(r"^name:\s*\S+", fm.group(1), re.M)
                     and re.search(r"^description:", fm.group(1), re.M))
        check("SKILL.md frontmatter 完整", fm_ok, "需要 name+description")
        if want:
            m = re.search(r"^>\s*版本:\s*(v[\d.]+)", sk, re.M)
            if m:
                check("SKILL.md 头部版本一致", m.group(1) == want,
                      f"SKILL.md={m.group(1)} skill.json={want}")

    # CHANGELOG
    cl_path = next((c for c in (ROOT / "versions" / "CHANGELOG.md", ROOT / "CHANGELOG.md") if c.exists()), None)
    if cl_path is not None and want:
        first = None
        for line in (text_of(cl_path) or "").splitlines():
            # 兼容表格行 `| v1.2.3 |` 与标题行 `## v1.2.3 (date)` / `# v1.2.3`
            m = re.match(r"^\|\s*(v[\d.]+)\s*\|", line) or re.match(r"^#{1,6}\s*(v[\d.]+)\b", line)
            if m:
                first = m.group(1)
                break
        check("CHANGELOG 最新版本一致", first == want, f"CHANGELOG={first} skill.json={want}")
    if cl_path is None and (cfg or {}).get("require_changelog"):
        check("CHANGELOG 存在", False, "缺失但 config.require_changelog=true")

    # README 标题版本
    rd_path = ROOT / "README.md"
    rd = text_of(rd_path) if rd_path.exists() else None
    if rd is not None and want:
        m = re.search(re.escape(name) + r"\s*[）)]?\s*v([\d.]+)", rd)
        got = ("v" + m.group(1)) if m else None
        check("README 标题版本一致", got == want, f"README={got} skill.json={want}")

    # scripts 语法
    scripts_dir = ROOT / "scripts"
    if scripts_dir.exists():
        for py in sorted(scripts_dir.glob("*.py")):
            r = subprocess.run([sys.executable, "-m", "py_compile", str(py)], capture_output=True)
            err = r.stderr.decode("utf-8", "ignore")[:200] if r.returncode else ""
            check(f"脚本语法 {py.name}", r.returncode == 0, err)

    # required_files
    for rel in (cfg or {}).get("required_files", []):
        check(f"资源存在 {rel}", (ROOT / rel).exists())

    # 敏感扫描
    leak_pat = re.compile(
        r"(ghp_[A-Za-z0-9]{20,}|GITHUB_TOKEN\s*=|password\s*[:=]\s*\S+|Authorization\s*[:=]\s*[Tt]oken\s+\S+)")
    targets = []
    if "--staged" in sys.argv:
        r = subprocess.run(["git", "diff", "--cached", "--name-only", "--diff-filter=ACM"],
                           capture_output=True, text=True)
        targets = [ROOT / n for n in r.stdout.splitlines()
                   if n and Path(n).suffix in (".md", ".json", ".py", ".yml", ".yaml", ".bat", ".ps1")]
    else:
        for pat in ("SKILL.md", "README.md", "skill.json", ".cicd/config.json"):
            p = ROOT / pat
            if p.exists():
                targets.append(p)
        if scripts_dir.exists():
            targets += [p for p in scripts_dir.glob("*.py") if p.name != "qa_checks.py"]
    leaks = []
    for p in targets:
        if not p.exists():
            continue
        txt = text_of(p)
        if txt is None:
            continue
        if leak_pat.search(txt):
            leaks.append(str(p.relative_to(ROOT)))
    check("无敏感信息泄漏", not leaks, "、".join(leaks) if leaks else "已扫描文本文件")

    print("-" * 40)
    if FAILED:
        print(f"QA FAILED: {len(FAILED)} 项未通过")
        sys.exit(1)
    print(f"QA PASSED — {name} v{ver or '?'} 质量门禁全绿")


if __name__ == "__main__":
    main()
