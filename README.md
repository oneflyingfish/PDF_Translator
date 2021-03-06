## 应用描述

基于`QT C++`和`Microsoft Office COM组件`编写的`PPT、Word、图片`批量转`PDF`程序

## 开发环境：

* 开发语言：`C++`  
* 开发环境：`QT 5.14.2 、 Windows 10 64位 + Visual Stdio 2019 Community + Office 2019 professional plus`   


## 运行环境：

使用环境：`仅安装了Microsoft Office 2007及以上版本的Windows`  

## 程序功能：

* 目标目录计算
    * 智能目录修正
    * 自动目录提示
* 自定义转化类别
    * 幻灯片
    * 文档
    * 图片
* 子目录
    * 遍历
    * 忽略
* 目标目录结构
    * 保持
    * 平铺
* 非目标文件
    * 忽略
    * 拷贝
* 进度显示
* 程序控制
    * 运行
    * 暂停

## 支持格式：

| 类别               |                支持格式                | 无内容文件 | 路径空格 |
| :----------------- | :------------------------------------: | :--------: | :------: |
| 幻灯片(PowerPoint) |           `*.ppt`、 `*.pptx`           |    跳过    |   支持   |
| 文档(Word)         |           `*.doc` 、`*.docx`           |    支持    |   支持   |
| 表格(Excel)        |     核心代码存在bug，暂被取消支持      |    ----    |  -----   |
| 图片               | `*.png` 、`*.jpg`、 `*.jpeg`、 `*.bmp` |    ----    |   支持   |

## 已通过测试环境：

|    windows版本    |            office版本            | 是否能运行 |
| :---------------: | :------------------------------: | :--------: |
| windows 10 专业版 | Microsoft Office 2019 专业增强版 |     √      |
| windows 10 家庭版 |       Microsoft Office 365       |     √      |
|       ----        |              -----               |   待测试   |

## 安装包下载
* [安装包](https://github.com/oneflyingfish/PDF_Translator/releases)
