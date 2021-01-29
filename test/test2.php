<?php

include(dirname(__DIR__).DIRECTORY_SEPARATOR.'vendor/autoload.php');


class UserSheet extends \Ms100\ExcelImport\Sheet
{
    /**
     * [
     *    字段key => [ '字段表头', '字段为空或不存在时的默认值', '单元格宽度']
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
$sheet2 = new UserSheet(__DIR__.'/test.xlsx', 'Sheet2');
var_dump($sheet2->readOne());

var_dump($sheet1->readMany(3));
var_dump($sheet2->readMany(3));









