<?php

namespace Tianmiao\Excel;

use Exception;
use ZipArchive;

class ExcelZipExport extends ExcelExport
{

    /**
     * @var mixed
     */
    private $zip_path;

    /**
     * @var array
     */
    private $head = [];

    /**
     * @var array
     */
    private $excel_path_arr = [];

    /**
     * 单文件一次最多导出数据最大行数，默认60000
     *
     * @var int
     */
    private $single_max_row = 60000;

    /**
     * 最多导出数据最大行数，默认600000
     *
     * @var int
     */
    private $max_num = 600000;

    /**
     * 单文件重命名后缀
     *
     * @var int
     */
    private $file_num = 1;

    /**
     * @param $template_excel_path
     * @param $option
     */
    public function __construct($template_excel_path = "", $option = [])
    {
        if (!empty($option['zip_path'])) {
            $this->zip_path = $option['zip_path'];
        }

        parent::__construct($template_excel_path, $option);
    }

    /**
     *
     * @param $zip_path
     */
    public function setZipPath($zip_path)
    {
        if (empty($zip_path)) {
            $this->setError("未传递保存目录");
        }

        $this->zip_path = $zip_path;
    }

    /**
     * 写入数据（当单个excel超出限制后，自动切换下一个excel）
     * @param  array  $data
     */
    public function write(array $data)
    {
        // 初始化写入表头
        parent::write($this->head);
        $max_num = 0;

        foreach ($data as $v) {
            $max_num++;

            if ($max_num < $this->max_num) {
                // 写入数据
                parent::write($v);

                if ($this->getRow() > $this->single_max_row) {
                    parent::save($this->getCurrentExcelPath());

                    //为下一个xls文件写入表头
                    parent::write($this->head);
                }
            } else {
                $this->setError("超出" . $this->max_num);
            }
        }
    }

    /**
     * 获取当前Excel路径
     *
     * @return string
     */
    private function getCurrentExcelPath()
    {
        $trimmed = str_replace(strrchr($this->zip_path, "."), "", $this->zip_path);

        $zip_path = $trimmed . '_' . $this->file_num++ . '.xls';

        $this->excel_path_arr[] = $zip_path;

        return $zip_path;
    }

    /**
     * 设置单个excel表头
     * @param  array  $head
     */
    public function setHead(array $head)
    {
        $this->head = $head;
    }

    /**
     * 保存并进行压缩
     * @param  string  $path
     * @return bool
     * @throws Exception
     */
    public function save($path = "")
    {
        parent::save($this->getCurrentExcelPath());

        return $this->createZip($this->excel_path_arr);
    }

    /**
     * @param $fileList
     * @return bool
     * @throws Exception
     */
    public function createZip($fileList)
    {
        $filePath = dirname($this->zip_path);

        $zipFileName = $filePath . DIRECTORY_SEPARATOR . basename($this->zip_path);

        //目录不存在则创建
        $this->buildDir($zipFileName);

        $zip = new ZipArchive();
        if (!$zip->open($zipFileName, $zip::CREATE)) {
            $this->setError("创建" . $zipFileName . "失败");

            return false;
        }

        $res = $this->addFileZip($filePath, $fileList, $zip);

        $zip->close();

        if ($res) {
            //清理文件
            $this->delFile($filePath, $fileList);
        }

        return true;
    }

    /**
     * 打包指定目录下，指定文件 zip
     *
     * @param string $filePath  打包的文件路径
     * @param array $fileList  需要打包的文件
     * @param ZipArchive $zip  ZipArchive对象
     * @return false|mixed
     * @throws Exception
     */
    function addFileZip($filePath, $fileList, $zip)
    {
        $res = false;

        $handler = opendir($filePath);

        try {
            foreach ($fileList as $value) {
                //文件加入zip对象
                $res = $zip->addFile($value, basename($value));
            }
        } catch (Exception $e) {
            $this->setError($e->getMessage());
        }

        closedir($handler);

        return $res;
    }

    /**
     * 目录不存在怎创建目录
     *
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
     * 删除文件
     *
     * @param $dir
     * @param  array  $fileList
     * @throws Exception
     */
    private function delFile($dir, $fileList)
    {
        try {
            if (is_dir($dir)) {
                foreach ($fileList as $value) {
                    unlink($value);
                }
            }
        } catch (Exception $e) {
            $this->setError($e->getMessage());
        }
    }
}