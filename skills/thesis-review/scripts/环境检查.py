# -*- coding: utf-8 -*-
"""
软工本科毕业论文指导老师 Skill — 环境检查脚本
运行方式：python 环境检查.py
"""

import sys
import importlib


def check_module(name, display_name=None):
    """检查 Python 包是否已安装"""
    if display_name is None:
        display_name = name
    try:
        mod = importlib.import_module(name)
        version = getattr(mod, "__version__", "已安装")
        return True, f"  [OK] {display_name}: {version}"
    except ImportError:
        return False, f"  [缺失] {display_name}: 未安装，请运行 pip install {name}"


def check_skill_file():
    """检查 SKILL.md 是否存在"""
    import os
    skill_path = os.path.expanduser("~/.workbuddy/skills/软工本科毕业论文指导老师/SKILL.md")
    if os.path.exists(skill_path):
        return True, f"  [OK] SKILL.md: {skill_path}"
    else:
        return False, f"  [未安装] SKILL.md 未找到，请将 SKILL.md 复制到 ~/.workbuddy/skills/软工本科毕业论文指导老师/"


def check_docx_plugin():
    """检查 docx 插件是否安装"""
    import os
    plugin_dir = os.path.expanduser(
        "~/.workbuddy/plugins/marketplaces/codebuddy-plugins-official/plugins/docx/"
    )
    if os.path.exists(plugin_dir):
        scripts = ["scripts/comment.py", "scripts/office/unpack.py", "scripts/office/pack.py"]
        missing = [s for s in scripts if not os.path.exists(os.path.join(plugin_dir, s))]
        if missing:
            return False, f"  [不完整] docx 插件已安装但缺文件: {missing}"
        return True, f"  [OK] docx 插件: {plugin_dir}"
    else:
        return False, f"  [未安装] docx 插件未安装，请在 WorkBuddy 插件市场搜索安装"


def main():
    print("=" * 60)
    print("软工本科毕业论文指导老师 Skill — 环境检查")
    print("=" * 60)
    print()

    # Python 版本
    ver = sys.version_info
    if ver.major >= 3 and ver.minor >= 9:
        print(f"  [OK] Python {ver.major}.{ver.minor}.{ver.micro}")
    else:
        print(f"  [警告] Python {ver.major}.{ver.minor}，建议 >= 3.9")
    print()

    # Python 包
    print("--- Python 包 ---")
    all_ok = True
    for mod, display in [
        ("docx", "python-docx"),
        ("pdfplumber", "pdfplumber"),
        ("lxml", "lxml"),
        ("defusedxml", "defusedxml"),
    ]:
        ok, msg = check_module(mod, display)
        print(msg)
        if not ok:
            all_ok = False
    print()

    # Skill 文件
    print("--- Skill 安装 ---")
    ok, msg = check_skill_file()
    print(msg)
    if not ok:
        all_ok = False
    print()

    # docx 插件
    print("--- docx 插件 ---")
    ok, msg = check_docx_plugin()
    print(msg)
    if not ok:
        all_ok = False
    print()

    # 总结
    print("=" * 60)
    if all_ok:
        print("全部通过！Skill 可以正常使用了。")
    else:
        print("有项目未通过，请按上方提示安装缺失的依赖。")
        print()
        print("快速安装命令：")
        print("  pip install python-docx pdfplumber lxml defusedxml")
        print("  然后在 WorkBuddy 插件市场安装 docx 插件")
        print("  最后将 SKILL.md 复制到 ~/.workbuddy/skills/软工本科毕业论文指导老师/")
    print("=" * 60)


if __name__ == "__main__":
    main()
