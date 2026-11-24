# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：FinSummaryTool 客户端。

构建（Windows / macOS 均适用，产物平台跟随构建机）：
    pyinstaller --noconfirm --clean FinSummaryTool.spec

产物：
    dist/FinSummaryTool/FinSummaryTool.exe   （Windows，onedir 单文件夹分发）
    dist/FinSummaryTool.app                  （macOS 应用包）

资源打包：
    mapping.json 与 summary_template.xlsx 打入 fin_summary/resources，
    客户端不依赖工作目录（双击启动时 CWD 任意）。
    日志/底稿写入用户主目录 ~/FinSummaryTool/logs（打包模式自动切换）。
"""
import sys

a = Analysis(
    ['client.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('fin_summary/resources/mapping.json', 'fin_summary/resources'),
        ('templates/summary_template.xlsx', 'fin_summary/resources'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 剔除本应用用不到的重量级 Qt 模块（PyInstaller 的 PySide6 hook 默认会多收集一批）
    excludes=[
        'PySide6.QtQml', 'PySide6.QtQuick', 'PySide6.QtQuick3D', 'PySide6.QtQuickWidgets',
        'PySide6.QtPdf', 'PySide6.QtPdfWidgets',
        'PySide6.QtVirtualKeyboard',
        'PySide6.QtNetwork', 'PySide6.QtDBus',
        'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtWebChannel',
        'PySide6.QtWebSockets', 'PySide6.QtWebViewer',
        'PySide6.QtCharts', 'PySide6.QtDataVisualization', 'PySide6.QtGraphs',
        'PySide6.Qt3DCore', 'PySide6.Qt3DRender',
        'PySide6.QtMultimedia', 'PySide6.QtMultimediaWidgets',
        'PySide6.QtSql', 'PySide6.QtTest', 'PySide6.QtDesigner',
        'PySide6.QtBluetooth', 'PySide6.QtNfc', 'PySide6.QtPositioning',
        'PySide6.QtSensors', 'PySide6.QtSerialPort', 'PySide6.QtTextToSpeech',
        'PySide6.QtStateMachine', 'PySide6.QtScxml', 'PySide6.QtHttpServer',
    ],
    noarchive=False,
)

# 再过滤二进制：本应用仅用 QtCore/QtGui/QtWidgets(+QtSvg/QtDBus/QtOpenGL)，
# 剔除 QtPdf/QtQml/QtQuick/QtNetwork 等框架（PySide6 hook 会整体收集 Qt/lib，需在此过滤），
# 以及依赖已剔除模块的插件（tls/网络状态/qml 调试/虚拟键盘/iOS 触摸）
_QT_DROP = ('QtPdf', 'QtQml', 'QtQuick', 'QtNetwork', 'QtVirtualKeyboard', 'QtQuick3D')
_BIN_DROP = ('tls/', 'networkinformation/', 'qmltooling/', 'virtualkeyboard',
             'qtuiotouchplugin', 'libqpdf')


def _drop(name: str) -> bool:
    low = name.lower()
    if any(p in low for p in _BIN_DROP):
        return True
    norm = name.replace('Qt6', 'Qt')  # Windows 上为 Qt6Pdf.dll 等
    return any(p in norm for p in _QT_DROP)


a.binaries = [b for b in a.binaries if not _drop(b[0])]
a.datas = [d for d in a.datas if not _drop(d[0])]
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='FinSummaryTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,  # GUI 程序，不显示控制台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='FinSummaryTool',
)

if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='FinSummaryTool.app',
        bundle_identifier='com.finsummarytool.app',
        version='2.17.0',
        info_plist={
            'CFBundleDisplayName': '月度财务数据汇总表生成工具',
            'CFBundleShortVersionString': '2.17.0',
            'NSHighResolutionCapable': True,
        },
    )
