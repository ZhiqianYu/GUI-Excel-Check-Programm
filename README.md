# Excel Data Matching and Filling Tool / Excel 数据匹配与填充工具

This project provides an easy-to-use interface that allows users to load Excel files, select different sheets and columns for matching and data filling. Based on the specified matching rules (such as exact match, contains, starts with, ends with, etc.), the data from one table can be filled into another. After filling, the user can choose to save the result either to the original file or a new one.

该项目提供了一个简单易用的界面，允许用户加载Excel文件，选择不同表格和列进行匹配与数据填充。可以根据指定的匹配规则（如完全匹配、包含、以...开头/结尾等）将一个表格的数据填充到另一个表格中。填充完成后，用户可以选择保存结果到原文件或新文件。

## Features / 功能概述

- **Select Excel File**: Allows the user to select an Excel file to be processed.
  
  **选择Excel文件**：用户可以选择一个Excel文件，工具会加载其中的所有表格。

- **Pair Matching**: 
  - Match data between two tables based on a selected column from each table.
  - Supports multiple matching rules: Exact match, Contains, Starts with, Ends with.

  **配对匹配**：
  - 从一个表格的指定列中查找与另一个表格的某列匹配的数据。
  - 支持多种匹配规则：完全匹配、包含、以...开头/结尾等。

- **Fill Matched Data**: 
  - After matching, fill the data into a selected column of another table.

  **填充匹配结果**：
  - 根据匹配结果，将数据填充到指定表格的指定列。

- **Save Result**: 
  - You can choose to save the filled result either in the original file or as a new file.

  **保存结果**：
  - 用户可以选择覆盖原文件或保存为新文件。

## Installation / 安装

1. Clone the repository:

    ```bash
    git clone https://github.com/yourusername/excel-matcher.git
    ```

2. Install the dependencies:

    ```bash
    pip install -r requirements.txt
    ```

## Usage / 使用方法

1. Start the application: （can also directly use the exe file, 也可直接使用exe程序）

    ```bash
    python GUI_Mobile_Check.py
    ```

2. The user interface will open. Follow these steps:

    用户界面会打开，按照以下步骤操作：

   - **Select Excel File**: Click the "Select File" button to choose the Excel file you want to process.
   
     **选择Excel文件**：点击“Select File”按钮，选择您需要处理的Excel文件。
   
   - **Select Tables and Columns**:
     - Select the table and column for pair matching.
     - Select the table and column for searching data.
     - Select the table and column to fill in the matched data.

     **选择表格与列**：
     - 选择您希望进行配对匹配的表格和列。
     - 选择您希望搜索数据的表格和列。
     - 选择您希望填充数据的表格和列。

   - **Choose Match Type**: Select the match type (Exact match, Contains, Starts with, Ends with).

     **选择匹配类型**：选择您需要的匹配规则（如完全匹配、包含、开始于、结束于）。

   - **Match & Fill**: Click the "Match & Fill" button to perform the matching and fill the results.

     **执行匹配与填充**：点击“Match & Fill”按钮，工具会根据选择的规则进行匹配，并将匹配结果填充到目标表格中。

3. After completion, you can choose to save the filled data. By default, it will overwrite the original file, or you can save it as a new file.

   完成后，您可以选择保存填充后的结果，默认会覆盖原文件，或另存为新文件。

## Code Structure / 代码结构

- `GUI_Mobile_Check.py`: The main PyQt program for the GUI.
  
  **`GUI_Mobile_Check.py`**：主要的PyQt界面程序。
  
- `requirements.txt`: The dependency file that lists the required libraries.
  
  **`requirements.txt`**：依赖库文件，列出了项目所需的Python库。
  
- `pandas`: Used for reading and manipulating Excel files.
  
  **`pandas`**：用于读取和操作Excel文件。
  
- `PyQt6`: Used for creating the graphical user interface.
  
  **`PyQt6`**：用于创建图形界面。
  
- `xlsxwriter`: Used for saving Excel files.
  
  **`xlsxwriter`**：用于保存Excel文件。

## Dependencies / 依赖

- `PyQt6`
- `pandas`
- `xlsxwriter`

Please make sure you have installed the above libraries, or you will not be able to run the program.

请确保您已安装上述库，否则您将无法运行程序。

## Example / 示例

Here is a simple usage example:

以下是一个简单的操作流程：

1. Select an Excel file that contains two sheets: `Sales` and `Employees`.
   
   选择一个Excel文件，文件包含两个表格：`Sales` 和 `Employees`。

2. In the "Choose Table for Pair Check", select the `Sales` table, and choose the column `EmployeeID` to pair with `EmployeeID` in the `Employees` table.
   
   在“Choose Table for Pair Check”中选择`Sales`表格，选择与`Employees`表格中的`EmployeeID`进行配对。

3. In the "Choose Column for Pair Check", select `EmployeeID`.
   
   在“Choose Column for Pair Check”中选择`EmployeeID`。

4. In the "Choose Column of Data for Fill In", select the column `Name`.
   
   在“Choose Column of Data for Fill In”中选择需要填充的列（如`Name`）。

5. After the matching is done, the `Sales` table's `Name` column will be filled with the matched data.

   执行匹配后，`Sales`表格中的`Name`列将被填充。

## Contribution / 贡献

Feel free to contribute to this project! If you have suggestions for improvements or want to fix bugs, please submit a pull request.

欢迎任何对本项目的贡献！如果您有改进建议或想要修复bug，请提交Pull Request。

## License / 许可证

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

此项目使用MIT许可证，详情请见 [LICENSE](LICENSE)。
