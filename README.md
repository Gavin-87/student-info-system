[学生信息管理系统-README.md](https://github.com/user-attachments/files/28187977/-README.md)
# 📋 学生信息管理系统 (Student Info Management System)

Python + Tkinter 桌面应用，基于 Excel 存储数据，实现学生信息的增删改查与成绩可视化统计。

## 功能特性

- **用户登录**：基于 Excel 存储的账号密码验证
- **信息管理**：学生信息的增加、删除、修改、查询
- **数据展示**：Treeview 表格展示所有学生数据，支持刷新
- **成绩统计**：Matplotlib 饼图展示数学/语文/英语各科成绩分布（优秀/良好/中等/及格/不及格）
- **学号搜索**：按学号精确查询学生信息

## 技术栈

- **语言**：Python
- **GUI**：Tkinter + sv_ttk（深色主题）
- **数据存储**：Excel（xlrd 读取 + xlutils 写入）
- **数据可视化**：Matplotlib
- **图片处理**：Pillow

## 项目结构

```
├── main.py           # 主程序（登录、增删改查、数据可视化）
├── students.xls      # 用户账号数据
└── grade.xls         # 学生成绩数据
```

## 如何运行

1. 安装依赖：

```bash
pip install xlrd xlutils matplotlib pillow sv_ttk
```

2. 确保 `students.xls` 和 `grade.xls` 与 `main.py` 在同一目录下

3. 运行：

```bash
python main.py
```

## 注意事项

- `students.xls` 和 `grade.xls` 为程序运行所需的数据文件，不含真实学生信息
- 登录账号密码存储在 `students.xls` 中

## 作者

GitHub: [Gavin-87](https://github.com/Gavin-87)
