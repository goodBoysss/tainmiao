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
     * @var array
     */
    private $column_width_arr = array();

    public function __construct($template_excel_path = "") {

        if (!empty($template_excel_path) && file_exists($template_excel_path)) {
            $this->load($template_excel_path);
        } else {
            $this->objPHPExcel = new Spreadsheet();
            $this->sheet = $this->objPHPExcel->setActiveSheetIndex(0);
        }

        $this->row = 1;

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
    public function load($template_excel_path) {
        //单元格不自动调整宽度
        $this->is_auto_column_width = 0;

        $this->objPHPExcel = IOFactory::load($template_excel_path);
        $this->sheet = $this->objPHPExcel->setActiveSheetIndex(0);
    }

    /**
     * 写入数据
     * @param array $data
     */
    public function write(array $data) {

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
    public function merge($start_cell, $end_cell) {
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
    public function center($cell){
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
     * 保存
     * @param $path
     * @return bool
     */
    public function save($path) {

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

            }else{

            }
            $objWriter = new Xls($this->objPHPExcel);


            $objWriter->save($path);

            //对象重新初始化
            $this->__construct();

        } catch (\Exception $e) {
            $result = false;

            $this->error = $e->getMessage();
        }

        return $result;

    }

    /**
     * 设置插入的起始行
     * @param int $row
     */
    public function setInsertRow($row = 1) {
        $this->row = $row;
    }


    /**
     * 获取错误消息
     * @return string
     */
    public function getError() {
        return $this->error;
    }

    /**
     * 设置自动适应宽度
     * @param  $char
     * @param  $v
     * @return null
     */
    private function setColumnWidth($char, $v) {

        $length = strlen($v);

        //长度预留
        $length = $length + 3;

        //自动适应列宽
        if (!empty($this->column_width_arr[$char])) {
            if ($this->column_width_arr[$char] < $length) {
                $this->column_width_arr[$char] = $length;
            }
        } else {
            $this->column_width_arr[$char] = $length;
        }
    }

    /**
     * 目录不存在怎创建目录
     * @param $path
     */
    private function buildDir($path) {
        $dir = dirname($path);
        if (!is_dir($dir)) {
            mkdir($dir, 0777, true);
        }
    }

    /**
     * 自动适应宽度
     */
    private function buildColumnWidth() {
        //调整宽度
        if (!empty($this->column_width_arr)) {
            foreach ($this->column_width_arr as $char => $column_width) {
                if ($column_width > $this->max_column_width) {
                    $column_width = $this->max_column_width;
                } elseif ($column_width < $this->min_column_width) {
                    $column_width = $this->min_column_width;
                }
                $this->sheet->getColumnDimension($char)->setWidth($column_width);
            }
        }
    }

    /**
     * 获取文件后缀
     * @param $file_name
     * @return string
     */
    private function getFileType($file_name) {
        $arr = explode('.', $file_name);

        $suffix = end($arr);

        return $suffix;
    }


}