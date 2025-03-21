# 数据可视化应用

一个使用Python Tkinter开发的数据可视化应用程序，支持：

- CSV数据导入和分析
- 多种图表类型：折线图、柱状图、散点图、饼图、热力图、直方图
- 数据清洗和异常值处理
- 自适应中文字体支持
- 暗黑/明亮主题切换
- 交互式数据浏览

## 功能特点

- 支持大数据集分页浏览
- 自动检测数值列进行可视化
- 智能图表类型推荐
- 支持图表保存和导出
- 数据搜索和过滤功能
- 数据库存储支持

## 使用方法

1. 加载CSV文件
2. 选择X轴和Y轴数据列
3. 选择图表类型
4. 创建可视化图表

## 支持的图表类型

- 折线图：适合显示随时间变化的趋势
- 柱状图：比较不同类别的数量
- 散点图：显示两个变量之间的关系
- 饼图：显示部分与整体的关系
- 热力图：通过颜色深浅展示数据分布
- 直方图：显示数据的分布情况

## 依赖库

- tkinter
- matplotlib
- pandas
- numpy
- sqlite3
- ttkthemes 