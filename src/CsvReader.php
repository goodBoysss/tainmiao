<?php
/**
 * CsvReader.php
 * ==============================================
 * Copy right 2015-2021  by https://www.tianmtech.com/
 * ----------------------------------------------
 * This is not a free software, without any authorization is not allowed to use and spread.
 * ==============================================
 * @desc : csv读取
 * @author: zhanglinxiao<zhanglinxiao@tianmtech.cn>
 * @date: 2021/11/26
 * @version: v1.0.0
 * @since: 2021/11/26 09:11
 */

namespace Tianmiao\Excel;
class CsvReader
{

    private $path = "";

    private $content = array();

    public function __construct($path) {
        $this->path = $path;
    }


    /**
     * 获取excel全部数据
     * @return array
     */
    public function getAllData() {
        $data = array();

        $path = $this->path;
        if (!empty($path) && file_exists($path)) {
            $content = file_get_contents($path);
            $content=iconv("gbk","utf-8//IGNORE",$content);
            $list = explode('\n', $content);
            foreach ($list as $info) {
                $data[] = explode(',', $info);
            }
        }


        return $data;
    }


}