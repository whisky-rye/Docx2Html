<?php

/**
 * Class Docx2Html
 * @desc 自封装超简单的doc文档转html网页类
 */
class Docx2Html{

    // 文件序号
    private $fileNo;
    // 应用域名
    private static $appDomain = "http://test.com";
    // 文件上传的根文件夹路径
    private static $uploadDir = "../static/uploads/";
    // 允许解释doc中的图片标签
    private static $imageTag = ["wpict","wdrawing"];
    // 允许读取doc中的图片后缀
    private static $imageExt = ["png","jpeg","jpg"];

    private $docPath;
    private $zipPath;
    private $unzipPath;
    private $rootPath;
    private $inImagePath;
    private $xmlPath;
    private $outImagePath;


    public function __construct($fileNo){
        $this->fileNo = $fileNo;
        $this->definePath();
        $this->doc2zip();
        $this->zip2unzip();
        $this->loadXmlToDom();
    }

    /**
     * 定义相关文件路径
     */
    private function definePath(){
        // 已经上传上来的doc文件路径
        $this->docPath = self::$uploadDir.'doc/'.$this->fileNo.".docx";
        // doc转压缩的zip文件路径
        $this->zipPath = self::$uploadDir.'zip/'.$this->fileNo.".zip";
        // zip解压缩后的文件夹路径
        $this->unzipPath = self::$uploadDir.'unzip/'.$this->fileNo."/";
        // 输出新的图片资源
        $this->inImagePath = self::$uploadDir."image/".$this->fileNo."/";
        @mkdir($this->inImagePath,0777);  // 建图片文件夹

        $this->rootPath = $this->unzipPath."word/";
   }


    private function doc2zip(){
        copy($this->docPath,$this->zipPath);
        if(!is_file($this->zipPath)){
            throw new Exception("Zip文件不存在");
        }
    }


    private function zip2unzip(){
        $zip = new ZipArchive();
        $rs = $zip->open($this->zipPath);
        if($rs !== true){
            throw new Exception("Zip解压失败");
        }
        $zip->extractTo($this->unzipPath);
        $zip->close();
    }

    /**
     * 从xml文件获取dom对象
     */
    private function loadXmlToDom(){
        $this->xmlPath = $this->rootPath."document.xml";
        // 原解压后的图片资源
        $this->outImagePath = $this->rootPath."media/image";

        // 找到xml文件
        $xml = file_get_contents($this->xmlPath);
        $xml = str_replace(":","",$xml);
        $handle = fopen($this->xmlPath,"w+");
        fwrite($handle,$xml);
        fclose($handle);

        // clearstatcache();
    }


    /**
     * 真正获取html代码
     */
    public function transferToHtml(){
        // 开始载入dom
        $doc = new DOMDocument();
        $doc->load($this->xmlPath);
        // unlink($this->xmlPath);
        /**
         * 纯文本
         * wdocument >> wbody >> wp >> wr >> wt >> "comtentText"
         * 非正常文本非黑色
         * wdocument >> wbody >> wp >> whyperlink >> wr >> wt >> "comtentText"
         * 图片
         * wdocument >> wbody >> wp >> wr >> wpict / wdrawing >> "to get imageSrc"
         */
        $wdocument = $doc->getElementsByTagName("wdocument")->item(0);
        // root tag
        $wbody = $wdocument->getElementsByTagName("wbody")->item(0);
        // body tag
        $wp = $wbody->getElementsByTagName("wp");
        // 行标签

        $wplen = $wp->length;
        $pngIndex = 1;  // 图片指针
        $html = '';

        for($i=0;$i<$wplen;$i++){  //  wp循环（一般为一段）

            $wr = $wp->item($i)->getElementsByTagName("wr");
            $wrlen = $wr->length;
            $html .= '<p>&nbsp;&nbsp;';

            for($s=0;$s<$wrlen;$s++){  //  wr循环（一般为一行）

                $wt = $wr->item($s)->getElementsByTagName("wt");
                $wtlen = $wt->length;

                if($wtlen == 0){  // 不存在文本标签（可能包含图片标签）

                    foreach( self::$imageTag as $tag ){
                        if($wr->item($s)->getElementsByTagName($tag)->length != 0){
                            foreach( self::$imageExt as $ext ){
                                if(file_exists($this->outImagePath.$pngIndex.".".$ext)){
                                    copy($this->outImagePath.$pngIndex.".".$ext,$this->inImagePath.$pngIndex.".".$ext);
                                    $html .= '<img src="'.self::$appDomain.'/'.$this->inImagePath.$pngIndex.".".$ext.'">';
                                    $html .= '<br><br>';
                                    $pngIndex++;
                                    break 2;
                                }
                            }	// for -- ext
                        }
                    } 		// for -- tag

                } // 判断图片标签结束
                else{		//  存在文本（含空文本）
                    $html .= $wt->item(0)->nodeValue;  //  每一行的主文本
                }

            }  // for -- wr

            /***  这里存疑根据具体情况判断是否使用  ***/
            /*
                // 判断是否含有非正常字体文字
                $whyperlink = $wp->item($i)->getElementsByTagName("whyperlink");
                // whyperlink是唯一标签
                if( $whyperlink->length != 0 ){
                    //  含有则输出
                    $html .= "<b>".$whyperlink->item(0)->getElementsByTagName("wr")->item(0)->getElementsByTagName("wt")->item(0)->nodeValue."</b>";
                }  // if exists whyperlink
            */

            $html .= '<p>';

        }  // for -- wp

        return $html;

    }


    /**
     * 把内容写入html文件
     */
    public function writeToHtml($html){
        // 暂时只有一种简单模板
        $tplPath = self::$uploadDir.'../tpl/single.html';
        $tplContent = file_get_contents($tplPath);
        $newContent = str_replace('{{{$toRepeatHtml}}}',$html,$tplContent);

        // 网页标题
        $title = strip_tags($html);
        $title = str_replace(' ',"",$title);
        $title = str_replace("&nbsp;","",$title);
        $title = mb_substr($title,0,8,"utf-8");
        $newContent = str_replace('{{{$toRepeatTitle}}}',$title,$newContent);

		$handle = fopen(self::$uploadDir.'html/'.$this->fileNo.'.html',"w+");
		fwrite($handle,$newContent);
		fclose($handle);

		return self::$appDomain.'/'.self::$uploadDir.'html/'.$this->fileNo.'.html';
    }


    /**
     * 清理转换过程中生成的所有临时文件
     */
    public function deleteTemp(){

        if(file_exists($this->zipPath)) {
            unlink($this->zipPath);
        }

        if(file_exists($this->docPath)) {
            unlink($this->docPath);
        }

        if(is_dir($this->unzipPath)) {
            $this->delDir($this->unzipPath);
        }

        clearstatcache();
    }


    /**
     * 删除一个目录下所有文件
     */
    private function delDir($dir) {
        // 先删除目录下的文件：
        $dh = opendir($dir);
        while($file = readdir($dh)){
            if($file != "." && $file != ".."){
                $fullpath = $dir."/".$file;
                if(!is_dir($fullpath)){
                    unlink($fullpath);
                } else {
                    $this->delDir($fullpath);
                }
            }
        }
        closedir($dh);
        // 删除当前文件夹：
        if(rmdir($dir)){
            return true;
        } else {
            return false;
        }
    }

}


