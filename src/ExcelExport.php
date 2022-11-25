<?php
/**
 * ExcelExport.php
 * ==============================================
 * Copy right 2015-2021  by https://www.tianmtech.com/
 * ----------------------------------------------
 * This is not a free software, without any authorization is not allowed to use and spread.
 * ==============================================
 * @desc : excel
 * @author: zhanglinxiao<zhanglinxiao@tianmtech.cn>
 * @date: 2021/11/11
 * @version: v1.0.0
 * @since: 2021/11/11 09:11
 */

namespace Tianmiao\Excel;

use mysql_xdevapi\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;

class ExcelExport
{

    /**
     * @var Spreadsheet
     */
    private $objPHPExcel;
    /**
     * @var Worksheet
     */
    private $sheet;

    /**
     * 当前插入行
     * @var int
     */
    private $row = 1;

    /**
     * 当前插入行（按工作区区分）
     * @var int
     */
    private $rows = array();

    /**
     * 错误提示
     * @var string
     */
    private $error = "";

    /**
     * 是否自动调整单元格宽度
     * @var int
     */
    private $is_auto_column_width = 1;

    /**
     * 列最小宽度
     * @var int
     */
    private $min_column_width = 10;


    /**
     * 列最大宽度
     * @var int
     */
    private $max_column_width = 50;


    /**
     * 列单元格宽度
     * @var string $template_excel_path 模板excel文件地址
     * @var array $option
     * @var array
     */
    private $column_width_arr = array();

    public function __construct($template_excel_path = "", $option = array())
    {
        $sheetIndex = 0;
        if (!empty($template_excel_path) && file_exists($template_excel_path)) {
            $this->load($template_excel_path);
        } else {
            unset($this->objPHPExcel);
            unset($this->sheet);

            $this->objPHPExcel = new Spreadsheet();
            $this->sheet = $this->objPHPExcel->setActiveSheetIndex($sheetIndex);
        }

        $this->row = 1;
        $this->rows = array(
            $sheetIndex => $this->row,
        );

        $this->error = "";

        $this->min_column_width = 10;

        $this->max_column_width = 50;

        $this->column_width_arr = array();
    }

    /**
     * 加载模板文件
     * @param $template_excel_path
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function load($template_excel_path)
    {
        //单元格不自动调整宽度
        $this->is_auto_column_width = 0;

        $this->objPHPExcel = IOFactory::load($template_excel_path);
        $this->sheet = $this->objPHPExcel->setActiveSheetIndex(0);
    }

    /**
     * 写入数据
     * @param array $data
     */
    public function write(array $data)
    {

        if (!empty($data[0]) && is_array($data[0])) {
            $data_arr = $data;
        } else {
            $data_arr = array($data);
        }
        foreach ($data_arr as $data) {
            if (!empty($data)) {
//                $ascii = 65;//A的ascii
                $col = 1;
                foreach ($data as $v) {
                    if ($col <= 26) {
                        $char = chr(64 + $col);
                    } else {

                        //出现多列时，例如：AA，AB

                        //取整
                        $n = floor($col / 26);
                        //余数
                        $r = $col % 26;

                        if ($r == 0) {
                            $n--;
                            $r = 26;
                        }

                        $char = chr(64 + $n) . chr(64 + $r);
                    }

                    if ($this->is_auto_column_width == 1) {
                        //前500行提取宽度
                        if ($this->row < 200) {
                            //设置自动适应宽度
                            $this->setColumnWidth($char, $v);
                        }
                    }

                    $this->sheet->setCellValue("{$char}{$this->row}", $v);


                    $col++;
                }

                $this->row++;
            }
        }
    }

    /**
     * 合并单元格
     * @param string $start_cell 例：A1
     * @param string $end_cell 例：G2
     * @return bool
     */
    public function merge($start_cell, $end_cell)
    {
        try {
            $this->sheet->mergeCells("{$start_cell}:{$end_cell}");
            $result = true;
        } catch (\Exception $e) {
            $this->error = $e->getMessage();
            $result = false;
        }

        return $result;
    }

    /**
     * 水平居中
     * @param string $cell 例：A1 或 A1:A2
     * @return bool
     */
    public function center($cell)
    {
        try {
            $this->sheet->getStyle($cell)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $result = true;
        } catch (\Exception $e) {
            $this->error = $e->getMessage();
            $result = false;
        }
        return $result;

    }


    /**
     * 新增工作区
     * @param string $title
     * @return bool
     */
    public function addSheet($title)
    {
        try {
            $oldSheetIndex = $this->objPHPExcel->getActiveSheetIndex();
            $newSheetIndex = count($this->rows);
            //切换工作区行数
            $this->switchRow($oldSheetIndex, $newSheetIndex);

            //新建工作区
            $clonedWorksheet = new Worksheet();
            $clonedWorksheet->setTitle($title);

            $this->sheet = $this->objPHPExcel->addSheet($clonedWorksheet);
            $this->objPHPExcel->setActiveSheetIndex($newSheetIndex);

            $result = true;
        } catch (\Throwable $e) {
            $this->error = $e->getMessage();
            $result = false;
        }
        return $result;
    }

    /**
     * 切换工作区
     * @param int $newSheetIndex
     * @return bool
     */
    public function switchSheet($newSheetIndex)
    {
        try {
            $oldSheetIndex = $this->objPHPExcel->getActiveSheetIndex();
            //切换工作区行数
            $this->switchRow($oldSheetIndex, $newSheetIndex);
            //获取工作区
            $this->objPHPExcel->setActiveSheetIndex($newSheetIndex);
            $this->sheet = $this->objPHPExcel->getActiveSheet();
            $result = true;
        } catch (\Throwable $e) {
            $this->error = $e->getMessage();
            $result = false;
        }
        return $result;
    }

    /**
     * 获取工作区
     * @return Worksheet
     */
    public function getSheet()
    {
        return $this->sheet;
    }

    /**
     * 获取excel对象
     * @return Spreadsheet
     */
    public function getExcel()
    {
        return $this->objPHPExcel;
    }

    /**
     * 保存
     * @param $path
     * @return bool
     */
    public function save($path = "")
    {

        $result = true;
        try {
            //目录不存在则创建
            $this->buildDir($path);

            //自动适应宽度
            $this->buildColumnWidth();

            //清除缓冲区,避免乱码
            //ob_end_clean();

            $suffix = $this->getFileType($path);
            if ($suffix == "xlsx") {

            } else {

            }

            //单元格默认居中
            $this->center("A:AC");

            //切换到第一个工作区
            $this->switchSheet(0);
            //生成文件
            $objWriter = new Xls($this->objPHPExcel);
            $objWriter->save($path);

            unset($objWriter);

            $this->objPHPExcel->__destruct();
            $this->sheet->__destruct();

            //对象重新初始化
            self::__construct();


        } catch (\Exception $e) {
            $result = false;
            $this->error = $e->getMessage();
            //执行失败抛错
            throw new Exception($this->error, '500');
        }

        return $result;

    }

    /**
     * 设置插入的起始行
     * @param int $row
     */
    public function setInsertRow($row = 1)
    {
        $this->row = $row;
    }


    /**
     * 获取错误消息
     * @return string
     */
    public function getError()
    {
        return $this->error;
    }

    /**
     * 切换工作区行数
     * @param $oldIndex
     * @param $newIndex
     */
    public function switchRow($oldIndex, $newIndex)
    {
        $this->rows[$oldIndex] = $this->row;
        if (!empty($this->rows[$newIndex])) {
            $this->setInsertRow($this->rows[$newIndex]);
        } else {
            $this->rows[$newIndex] = 1;
            $this->setInsertRow(1);
        }
    }

    /**
     * 设置自动适应宽度
     * @param  $char
     * @param  $v
     * @return null
     */
    private function setColumnWidth($char, $v)
    {

        $length = strlen($v);

        //长度预留
        $length = $length + 3;

        //工作区index
        $sheetIndex = $this->objPHPExcel->getActiveSheetIndex();
        if (!isset($this->column_width_arr[$sheetIndex])) {
            $this->column_width_arr[$sheetIndex] = array();
        }

        //自动适应列宽
        if (!empty($this->column_width_arr[$sheetIndex][$char])) {
            if ($this->column_width_arr[$sheetIndex][$char] < $length) {
                $this->column_width_arr[$sheetIndex][$char] = $length;
            }
        } else {
            $this->column_width_arr[$sheetIndex][$char] = $length;
        }
    }

    /**
     * 目录不存在怎创建目录
     * @param $path
     */
    private function buildDir($path)
    {
        $dir = dirname($path);
        if (!is_dir($dir)) {
            mkdir($dir, 0777, true);
        }
    }

    /**
     * 自动适应宽度
     */
    private function buildColumnWidth()
    {
        //调整宽度
        if (!empty($this->column_width_arr)) {

            foreach ($this->column_width_arr as $sheelIndex => $columnWidthArr) {
                $this->switchSheet($sheelIndex);
                foreach ($columnWidthArr as $char => $column_width) {
                    if ($column_width > $this->max_column_width) {
                        $column_width = $this->max_column_width;
                    } elseif ($column_width < $this->min_column_width) {
                        $column_width = $this->min_column_width;
                    }
                    $this->sheet->getColumnDimension($char)->setWidth($column_width);
                }
            }
        }
    }

    /**
     * 获取文件后缀
     * @param $file_name
     * @return string
     */
    private function getFileType($file_name)
    {
        $arr = explode('.', $file_name);

        $suffix = end($arr);

        return $suffix;
    }

    /**
     * 设置单元格颜色
     * @param string $cell 例：A1 或 A1:A2
     * @param string $color 颜色，例如：fc5531
     * @return bool
     */
    public function setColor($cell, $color)
    {
        $this->sheet->getStyle("$cell")->getFont()->getColor()->setRGB($color);
    }

    /**
     * 获取行数
     *
     * @return int
     */
    public function getRow()
    {
        return $this->row;
    }

    /**
     * @param $error
     * @return false
     */
    public function setError($error)
    {
        $this->error = $error;
        return false;
    }
}