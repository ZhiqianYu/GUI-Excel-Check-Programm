Excel 数据匹配与填充工具
此项目提供了一个简单易用的界面，允许用户加载Excel文件，选择不同表格和列进行匹配与数据填充。可以根据指定的匹配规则（如完全匹配、包含、以...开头/结尾等）将一个表格的数据填充到另一个表格中。填充完成后，用户可以选择保存结果到原文件或新文件。

功能概述
选择要处理的Excel文件。
配对匹配：
从一个表格的指定列中查找与另一个表格的某列匹配的数据。
支持多种匹配规则（完全匹配、包含、以...开头/结尾等）。
填充匹配结果：
根据匹配结果，将数据填充到指定表格的指定列。
可选择覆盖原文件或保存为新文件。
安装
克隆该仓库到本地：

bash
Kopieren
Bearbeiten
git clone https://github.com/yourusername/excel-matcher.git
安装依赖：

bash
Kopieren
Bearbeiten
pip install -r requirements.txt
使用方法
启动应用程序：

bash
Kopieren
Bearbeiten
python GUI_Mobile_Check.py
用户界面会打开，按照以下步骤操作：

选择Excel文件：点击“Select File”按钮，选择您需要处理的Excel文件。
选择表格与列：
选择您希望进行配对匹配的表格和列。
选择您希望搜索数据的表格和列。
选择您希望填充数据的表格和列。
选择匹配类型：选择您需要的匹配规则（如完全匹配、包含、开始于、结束于）。
执行匹配与填充：点击“Match & Fill”按钮，工具会根据选择的规则进行匹配，并将匹配结果填充到目标表格中。
完成后，您可以选择保存填充后的结果，默认会覆盖原文件，或另存为新文件。

代码结构
GUI_Mobile_Check.py：主要的PyQt界面程序。
requirements.txt：依赖库文件，列出了项目所需的Python库。
pandas：用于读取和操作Excel文件。
PyQt6：用于创建图形界面。
xlsxwriter：用于保存Excel文件。
依赖
PyQt6
pandas
xlsxwriter
请确保您已安装上述库，否则您将无法运行程序。

示例
以下是一个简单的操作流程：

选择一个Excel文件，文件包含两个表格：Sales 和 Employees。
在“Choose Table for Pair Check”中选择Sales表格，选择与Employees表格中的EmployeeID进行配对。
在“Choose Column for Pair Check”中选择EmployeeID。
在“Choose Column of Data for Fill In”中选择需要填充的列（如Name）。
执行匹配后，Sales表格中的Name列将被填充。
贡献
欢迎任何对本项目的贡献！如果您有改进建议或想要修复bug，请提交Pull Request。

许可证
此项目使用MIT许可证，详情请见 LICENSE。
