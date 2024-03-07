# README

## 1.文件目录说明

<img src="https://pic.imgdb.cn/item/65e972469f345e8d03c860a1.png" alt="image-20240307142933692.png">

#### tips
1.source>dynamic是问卷星生成的两个文件所在的文件夹
2.source>end是病例和测评所在的文件夹

> 1,2中的两个文件夹需要每次手动导入，不需要手动删除，程序运行完毕后会将其清除以便下一次使用

3.source>Static是静态文件夹，不需要用户进行操作

4.init.py是初始化代码，需要第一次在电脑上使用该程序的用户运行

4.run.py是功能代码，用以生成完整报告，运行完成后生成“心理报告.pdf”（生成的文件在PYTHON目录下，此处未显示）

## 2.使用说明

### *step0 程序初始化

对于**第一次**在电脑上运行本程序的用户需要先运行init.py来进行程序环境配置。

### step1 文件准备

1.将问卷星得到的文件放入source>dynamic文件夹

2.将病例图片和心理测评报告放入source>end文件夹

### step2 运行代码

**运行run.py**

source>dynamic和source>end内的文件已经被删除，”心理报告.pdf“生成完毕。