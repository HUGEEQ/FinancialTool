"""客户端可执行文件入口（PyInstaller 打包用）。

开发调试仍用 python -m fin_summary.gui；本文件供 PyInstaller 作为启动脚本，
使用绝对导入避免打包后相对导入失效。

--selftest：打包后自检（验证 mapping.json 与输出模板在客户端内可正常解析），不启动界面。
"""
import sys

from fin_summary.gui import main


def selftest() -> int:
    """打包自检：计算规则配置、项目名称基准（模板）、日志目录可写性。"""
    from fin_summary.config_loader import load_mapping
    from fin_summary.exporter import display_names, template_path
    from fin_summary.logger import RunLogger, default_log_root, gen_run_id

    cfg, items, version, sha = load_mapping()
    assert len(items) == 70, f"映射项应为 70，实际 {len(items)}"
    names = display_names()
    assert len(names) == 70, f"项目名称应为 70，实际 {len(names)}"

    rl = RunLogger(gen_run_id(), "0000-00")  # 日志目录可写性（打包模式写入用户主目录）
    rl.write_summary({"status": "SELFTEST"})
    assert rl.calc_path.exists()

    print(f"SELFTEST OK  version={version}  items={len(items)}  names={len(names)}")
    print(f"template={template_path()}")
    print(f"logroot={default_log_root()}")
    return 0


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        sys.exit(selftest())
    sys.exit(main())
