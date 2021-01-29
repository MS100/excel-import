<?php

namespace Ms100\ExcelImport;

use Vtiful\Kernel\Excel;

abstract class Sheet
{
    /**
     * @var Excel
     */
    private $excel;

    /**
     * @var array
     */
    private $column_key_indexes = [];

    /**
     * @var array
     */
    private $preprocess_methods = [];

    /**
     * @var array
     */
    private $required_columns = [];

    protected const LAYOUT = [
        //'字段key' => [ '匹配表头字段的正则表达式', '字段不存在时或预处理后为空值时，是否舍弃此行数据', '列的读取类型']
    ];
    protected const PREPROCESS_METHOD_PREFIX = 'preprocess';

    protected const COLUMN_PATTERN_KEY = 0;
    protected const COLUMN_REQUIRED_KEY = 1;
    protected const COLUMN_TYPE_KEY = 2;

    /**
     * Sheet constructor.
     *
     * @param string $filepath 文件路径
     * @param string $sheet_name sheet名，不传递取出第一个sheet
     *
     * @throws \Exception
     */
    final public function __construct(
        string $filepath,
        string $sheet_name = null
    ) {
        $config = [
            'path' => dirname($filepath),
        ];
        $this->excel = new Excel($config);

        $this->excel->openFile(basename($filepath));

        $sheet_list = $this->excel->sheetList();

        if (!is_null($sheet_name)) {
            if (!in_array($sheet_name, $sheet_list)) {
                throw new \Exception('未找到目标sheet', -1);
            }
        } else {
            $sheet_name = array_shift($sheet_list);
        }

        $this->excel->openSheet($sheet_name, Excel::SKIP_EMPTY_ROW);
        $this->setColumnKeyIndexes();
        $this->setColumnsType();
        $this->setColumnPreprocessMethods();
    }

    private function setColumnKeyIndexes()
    {
        $num = 100;
        $mark = false;
        $layout = static::LAYOUT;
        while ($num--) {
            $row = $this->excel->nextRow();
            foreach ($row as $index => $cell) {
                foreach ($layout as $key => $config) {
                    if (preg_match($config[self::COLUMN_PATTERN_KEY], $cell)) {
                        $this->column_key_indexes[$key] = $index;
                        unset($layout[$key]);
                        $mark = true;
                        break;
                    }
                }
                if (count($layout) == 0) {
                    break;
                }
            }
            if ($mark) {
                break;
            }
        }

        foreach ($layout as $item) {
            if ($item[self::COLUMN_REQUIRED_KEY]) {
                $this->column_key_indexes = [];
            }
        }
    }

    private function setColumnsType()
    {
        $type = [];
        foreach ($this->column_key_indexes as $key => $index) {
            $type[$index] = static::LAYOUT[$key][self::COLUMN_TYPE_KEY]
                ?? Excel::TYPE_STRING;
            empty(static::LAYOUT[$key][self::COLUMN_REQUIRED_KEY])
            || $this->required_columns[] = $key;
        }
        $this->excel->setType($type);
    }

    private function setColumnPreprocessMethods()
    {
        foreach ($this->column_key_indexes as $key => $index) {
            $method_name = self::PREPROCESS_METHOD_PREFIX
                .$this->toBigCamelCase($key);
            if (method_exists($this, $method_name)) {
                $this->preprocess_methods[$key] = $method_name;
            }
        }
    }

    private function convertData(array $row)
    {
        $data = [];
        foreach ($this->column_key_indexes as $key => $index) {
            $data[$key] = isset($this->preprocess_methods[$key])
                ? $this->{$this->preprocess_methods[$key]}($row[$index])
                : $row[$index];
        }

        return $data;
    }

    /**
     * 读取一行
     * @return array|null
     */
    final public function readOne()
    {
        if (empty($this->column_key_indexes)) {
            return null;
        }

        $row = $this->excel->nextRow();

        if (is_array($row)) {
            $data = $this->convertData($row);
            if ($this->isValid($data)) {
                return $data;
            } else {
                return $this->readOne();
            }
        } else {
            return null;
        }
    }

    /**
     * 读取多行
     * @param int $row_num
     *
     * @return array
     */
    final public function readMany(int $row_num)
    {
        if (empty($this->column_key_indexes)) {
            return [];
        }

        $row_num = min(1000, max(1, $row_num));
        $data = [];
        while ($row_num-- > 0) {
            $row = $this->readOne();

            if (is_array($row)) {
                $data[] = $row;
            } else {
                break;
            }
        }

        return $data;
    }

    /**
     * 读取剩余所有行
     *
     * @return array
     */
    final public function readRest()
    {
        if (empty($this->column_key_indexes)) {
            return [];
        }

        $rows = $this->excel->getSheetData();

        foreach ($rows as $index => $row) {
            $data = $this->convertData($row);
            if ($this->isValid($data)) {
                $rows[$index] = $data;
            } else {
                unset($rows[$index]);
            }
        }

        return array_values($rows);
    }

    private function toBigCamelCase(string $str)
    {
        $str = ucwords(strtolower($str), '_');

        $str = strtr($str, ['_' => '']);

        return $str;
    }

    private function isValid(array $row)
    {
        if (empty($row = array_filter($row))) {
            return false;
        }

        foreach ($this->required_columns as $key) {
            if (!isset($row[$key])) {
                return false;
            }
        }

        return true;
    }
}

