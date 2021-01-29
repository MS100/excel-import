<?php

namespace Ms100\ExcelImport;

class ExcelImport
{
    private $filepath = '';

    /**
     * ExcelImport constructor.
     *
     * @param string $filepath
     */
    final public function __construct(string $filepath)
    {
        $this->filepath = $filepath;
    }

    /**
     * @param string $sheet_class_name
     * @param string $sheet_name
     *
     * @return Sheet
     */
    final public function getSheet(string $sheet_class_name, string $sheet_name = null)
    {
        if(is_subclass_of($sheet_class_name,Sheet::class)){
            return new $sheet_class_name($this->filepath, $sheet_name);
        }
    }
}