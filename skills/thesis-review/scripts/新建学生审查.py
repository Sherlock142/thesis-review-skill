"""
快速初始化一个学生的审查目录结构
用法：python scripts/新建学生审查.py <学生姓名或学号>

示例：
  python scripts/新建学生审查.py 张三
  → 创建 01_学生论文/张三/、02_审查报告/张三/、03_整改论文/张三/
"""

import os
import sys

WORKSPACE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    if len(sys.argv) < 2:
        print("用法: python scripts/新建学生审查.py <学生姓名或学号>")
        print("示例: python scripts/新建学生审查.py 张三")
        sys.exit(1)

    student = sys.argv[1]
    dirs = [
        os.path.join(WORKSPACE, "01_学生论文", student),
        os.path.join(WORKSPACE, "02_审查报告", student),
        os.path.join(WORKSPACE, "03_整改论文", student),
    ]

    for d in dirs:
        os.makedirs(d, exist_ok=True)
        print(f"  ✓ {d}")

    print(f"\n学生 '{student}' 目录已就绪。")
    print(f"下一步：把论文文件放到 01_学生论文/{student}/ 中。")


if __name__ == "__main__":
    main()
