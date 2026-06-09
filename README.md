# 按号码提取数据

这是一个 Windows 图形界面工具，用于从多个期次的 Excel 文件中按每期配置的蓝球号码提取数据，并合并输出为新的 `.xlsx` 结果文件。

## 主要功能

- 选择包含多个 `.xlsx` 数据文件的文件夹
- 扫描每个文件中的期次
- 为每个期次设置 `1-16` 的筛选号码
- 支持生成号码配置模板，再导入配置 Excel
- 支持在界面表格的“筛选号码”列直接下拉选择
- 输出列为：`期号、专家名称、玩法、本期推荐、筛选号码`
- 可用 PyInstaller 打包为单文件 exe

## 开发环境

```powershell
python -m pip install -r requirements.txt
```

## 运行测试

```powershell
pytest tests -v
```

## 启动程序源码

```powershell
$env:PYTHONPATH='src'
python src\run_number_based_extractor.py
```

如果本机 Python 的 Tcl/Tk 路径异常，可以参考 `build_number_based_extractor_exe.ps1` 中的 `PYTHONHOME`、`TCL_LIBRARY`、`TK_LIBRARY` 设置。

## 打包 exe

```powershell
powershell -ExecutionPolicy Bypass -File .\build_number_based_extractor_exe.ps1
```

打包结果默认生成到 `outputs/` 目录。该目录是生成产物，不提交到 Git。

## 交接阅读顺序

如果由另一个 agent 继续开发，建议按下面顺序阅读：

1. `docs/superpowers/specs/2026-06-09-number-based-extraction-design.md`：当前最终需求和业务规则。
2. `docs/superpowers/plans/2026-06-09-number-based-extraction.md`：当前实现交接记录，包含已取消的旧需求说明。
3. `src/number_based_extractor/core.py`：核心筛选规则。
4. `src/number_based_extractor/excel_io.py`：Excel 读取、模板导入导出、结果保存。
5. `src/number_based_extractor/gui.py`：图形界面。
6. `tests/test_number_based_core.py` 和 `tests/test_number_based_excel_io.py`：当前行为验证。
