# XiaoFeng.Excel


![fayelf](https://user-images.githubusercontent.com/16105174/197918392-29d40971-a8a2-4be4-ac17-323f1d0bed82.png)

![GitHub top language](https://img.shields.io/github/languages/top/zhuovi/xiaofeng.excel?logo=github)
![GitHub License](https://img.shields.io/github/license/zhuovi/xiaofeng.excel?logo=github)
![Nuget Downloads](https://img.shields.io/nuget/dt/xiaofeng.excel?logo=nuget)
![Nuget](https://img.shields.io/nuget/v/xiaofeng.excel?logo=nuget)
![Nuget (with prereleases)](https://img.shields.io/nuget/vpre/xiaofeng.excel?label=dev%20nuget&logo=nuget)

Nuget：XiaoFeng.Excel

QQ群号：748408911 

QQ群二维码： 

![QQ 群](https://user-images.githubusercontent.com/16105174/198058269-0ea5928c-a2fc-4049-86da-cca2249229ae.png)

公众号： 

![畅聊了个科技](https://user-images.githubusercontent.com/16105174/198059698-adbf29c3-60c2-4c76-b894-21793b40cf34.jpg)

源码： https://github.com/zhuovi/XiaoFeng.Excel

教程： https://www.yuque.com/fayelf/xiaofeng


XiaoFeng.Excel操作，创建excel,编辑excel,读取excel内容，边框，字体，样式等功能。

## XiaoFeng.Excel

XiaoFeng.Excel generator with [XiaoFeng.Excel](https://github.com/zhuovi/XiaoFeng.Excel).

## Install

.NET CLI

```
$ dotnet add package XiaoFeng.Excel --version 1.0.0
```

Package Manager

```
PM> Install-Package XiaoFeng.Excel --Version 1.0.0
```

PackageReference

```
<PackageReference Include="XiaoFeng.Excel" Version="1.0.0" />
```

Paket CLI

```
> paket add XiaoFeng.Excel --version 1.0.0
```

Script & Interactive

```
> #r "nuget: XiaoFeng.Excel, 1.0.0"
```

Cake

```
// Install XiaoFeng.Excel as a Cake Addin
#addin nuget:?package=XiaoFeng.Excel&version=12.0.0

// Install XiaoFeng.Excel as a Cake Tool
#tool nuget:?package=XiaoFeng.Excel&version=1.0.0
```

# XiaoFeng.Excel 操作


```csharp

var excel = new XiaoFeng.Excel.Workbook();
//打开excel
excel.Open(@"E://a.xlsx");
///excel中所有的表
var sheets = excel.Sheets;
//第一张表中所有行
var rows = excel.Sheets.FirstOrDefault().Rows;
//第一张表中第一行所有列
var cells = rows.FirstOrDefault().Cells;
//单元格值
var cellValue = cell.Value;
//单元格位置
var location = cell.Location;
//单元格格式
var format = cell.Format;
//单元格所在行
var cellRow = cell.Row;
//单元格所在列
var cellColumn = cell.Column;

```

# 作者介绍

* 网址 : http://www.fayelf.com
* QQ : 7092734
* EMail : jacky@fayelf.com