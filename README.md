# 自动化排座管理系统

基于 Python 和 CustomTkinter 的 Windows 桌面排座工具。支持导入人员名单、自定义场地座位表、自动排座、座位预览、名单搜索及 Excel 导出，所有人员和场地数据均在本机处理。

## 使用

1. 从[版本发布页](https://github.com/TheWestofTokyo/Automated-Offline-Seating-Arrangement-System-Chinese-version/releases/latest)下载 Windows 程序并运行。
2. 默认场地为 **大学生活动中心礼堂**，含 532 个位置，其中普通席 498 个。
3. 如需其他场地，点击“导入新座位表”，选择 Excel 工作表。
4. 导入人员名单，核对记录和提示，选择“前排优先”或“同班集中”后开始排座。
5. 点击座位查看人员信息，或在人员明细中搜索。导出将生成座位图、个人明细和班级座位分配表。

新场地会保存到程序同目录的 `场地库` 文件夹，重启后可再次选择。切换场地会清除旧排座结果并保留人员名单，需重新排座。

## 座位表

模板位于 [templates](templates)，支持以下两种结构化 `.xlsx` 格式：

| 格式 | 字段 |
| --- | --- |
| 逐座明细 | 场地名称、排号、区域、座位号、用途 |
| 连续座位段 | 场地名称、排号、区域、起始座位号、结束座位号、用途 |

用途支持普通、主席团、老师、桌子、禁排。场地名称可省略，省略时取文件名。每个工作表只描述一个场地，排号和座位号为正整数，座位不能重复或范围重叠。

本程序不直接识别图片、PDF或任意合并单元格绘制的座位图；请按模板整理。

## 人员名单

必需列为“姓名”和“班级”（可用“单位”），身份列可选。主席团和老师使用对应预留席位，可不填写班级。重复记录会提醒并保留，不会自动删除。

## 从源码运行

建议使用 64 位 Python 3.12：

```powershell
python -m pip install -r requirements.txt
python app.py
```

回归测试：

```powershell
python -m unittest discover -p "test_*.py" -v
```

Windows 打包命令见 [build.ps1](build.ps1)。

主分支的程序文件更新后，GitHub Actions 会运行测试并创建 Windows 版本发布。发布页提供单文件 EXE、附带模板的压缩包及 SHA256 校验值。

## 代码结构

- `app.py`：桌面界面与后台任务。
- `venue.py`：座位表导入、校验及场地库保存。
- `seating_core.py`：内置礼堂、人员解析、排座与 Excel 导出。
- `test_seating.py`、`test_venue.py`：29 项回归测试，含两种策略共 200 组随机名单。
- `LICENSES`：第三方依赖许可。

完整操作规则见 [使用说明](使用说明.txt)。当前版本未包含手动换座、撤销或排座项目存档；场地库只保存布局，不保存人员名单和排座结果。

界面使用 [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)。
