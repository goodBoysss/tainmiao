<?php
/**
 * ExcelRead.php
 * ==============================================
 * Copy right 2015-2021  by https://www.tianmtech.com/
 * ----------------------------------------------
 * This is not a free software, without any authorization is not allowed to use and spread.
 * ==============================================
 * @desc : excel读取
 * @author: zhanglinxiao<zhanglinxiao@tianmtech.cn>
 * @date: 2021/11/26
 * @version: v1.0.0
 * @since: 2021/11/26 09:11
 */

namespace Tianmiao\Excel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ExcelReader
{

    /**
     * @var Spreadsheet
     */
    private $objPHPExcel;
    /**
     * @var Worksheet
     */
    private $sheet;

    public function __construct($excel_path)
    {
        $inputFileType = IOFactory::identify($excel_path);

        if (empty($inputFileType) || !in_array($inputFileType, array('Xlsx', 'Xls'))) {
            throw new \Exception("文件类型不合法");
        }

        $objReader = IOFactory::createReader($inputFileType);
        $this->objPHPExcel = @$objReader->load($excel_path);
        $this->sheet = $this->objPHPExcel->setActiveSheetIndex(0);

    }

    /**
     * 获取行数
     * @return int
     */
    public function getRowCount()
    {
        return $this->sheet->getHighestRow();
    }

    /**
     * 获取excel全部数据
     * @return array
     */
    public function getAllData()
    {
        $data = array();
        try {
            $highest_row = $this->sheet->getHighestRow();
            $highest_column = $this->sheet->getHighestColumn();


            for ($i = 1; $i <= $highest_row; $i++) {
                for ($j = 'A'; $j <= $highest_column; $j++) {
                    $value = $this->sheet->getCell("{$j}{$i}")->getValue();
                    $data[$i - 1][] = $value;
                }
            }
        } catch (\Exception $e) {

        }

        return $data;
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


}