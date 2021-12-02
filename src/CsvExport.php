<?php
/**
 * CsvExport.php
 * ==============================================
 * Copy right 2015-2021  by https://www.tianmtech.com/
 * ----------------------------------------------
 * This is not a free software, without any authorization is not allowed to use and spread.
 * ==============================================
 * @desc : csv文件导出
 * @author: zhanglinxiao<zhanglinxiao@tianmtech.cn>
 * @date: 2021/11/11
 * @version: v1.0.0
 * @since: 2021/11/11 09:11
 */

namespace Tianmiao\Excel;
class CsvExport
{

    private $data = array();

    private $error = "";

    /**
     * 当前插入行
     * @var int
     */
    private $row = 1;


    public function __construct() {

        $this->error = "";
        $this->data = array();
        $this->row = 1;

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
                $this->data[$this->row] = array();
                foreach ($data as $v) {
                    $this->data[$this->row][] = $v;
                }

                $this->row++;
            }
        }
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

            $w = fopen($path, "w+");

            if (!empty($this->data)) {
                foreach ($this->data as $list) {
                    fputcsv($w, $list);
                }
            }


            fclose($w);
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
     * 目录不存在怎创建目录
     * @param $path
     */
    private function buildDir($path) {
        $dir = dirname($path);
        if (!is_dir($dir)) {
            mkdir($dir, 0777, true);
        }
    }


}