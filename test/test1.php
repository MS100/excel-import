<?php

include(dirname(__DIR__).DIRECTORY_SEPARATOR.'vendor/autoload.php');


class UserSheet extends \Ms100\ExcelImport\Sheet
{
    /**
     * [
     *    字段key => [ '匹配表头字段的正则表达式', '字段不存在时或预处理后为空值时，舍弃此行数据', '读取单元格类型']
     * ]
     */
    protected const LAYOUT
        = [
            'name'   => ['#姓名#', true, \Vtiful\Kernel\Excel::TYPE_STRING,],
            'age'    => ['#年龄#', false, \Vtiful\Kernel\Excel::TYPE_INT],
            'height' => ['#身高#', false, \Vtiful\Kernel\Excel::TYPE_INT],
            'birth'  => ['#生日#', false, \Vtiful\Kernel\Excel::TYPE_TIMESTAMP],
        ];

    protected function preprocessName($name)
    {
        return $name == '无名氏' ? '' : $name;
    }

    protected function preprocessHeight($height)
    {
        return is_int($height) ? $height : 0;
    }

    protected function preprocessBirth($birth)
    {
        return is_int($birth) ? $birth : intval(strtotime($birth));
    }
}

$sheet1 = new UserSheet(__DIR__.'/test.xlsx');
var_dump($sheet1->readOne());
var_dump($sheet1->readMany(2));
var_dump($sheet1->readRest());






