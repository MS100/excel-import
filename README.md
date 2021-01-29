## 说明
基于php扩展 xlswriter，[手册](https://xlswriter-docs.viest.me/) 和 [Pecl](https://pecl.php.net/package/xlswriter)。

提供对excel表格快速导入。

## 安装
```bash
composer require ms100/excel-import
```

## 使用

#### 导入一个sheet
[见代码](test/test1.php)

 
#### 导入多个sheet

[见代码](test/test2.php)

#### 注意事项 


* 所有的读取数据方法都是【游标读取】，它们影响同一个游标；即第一次调用 `\Ms100\ExcelImport\Sheet::readOne` 读第一行，再调用任何读取方法都将从第二行开始读取。

* 自定义的preprocess方法可以做单元格数据的预处理。

* 读取数据时如果【单元格数据】符合所 LAYOUT 常量所设置的【列的读取类型】，那么会读取出正确的【类型值】；否则不能得到正确的【类型值】。
    * 例如：设置【列的读取类型】为 `\Vtiful\Kernel\Excel::TYPE_TIMESTAMP`，但是表中实际是写了个字符串，则得不到正确的时间戳；这时，用preprocess方法，再处理一下会很方便。
