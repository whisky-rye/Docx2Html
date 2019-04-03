- php操作docx文档转html网页

- 转换原理是将docx文件转换成zip文件，再解压到文件夹，读取里面的/word/document.xml的xml树，逐一解析标签获取原文字到html字符，同时将/word/media里的文档图片复制到目标目录，并插入图片路径到html字符对应位置，最后把html字符塞进预先的模板，生成html文件。

- 暂时仅支持.docx的文档，.doc的解压缩包目录结构很不一样，无法兼容处理。如果想体验先用WPS或者Office转为.docx再使用。

- 目前只支持将docx的文字和图片提取到网页，不支持其他高级word功能。

- static目录说明：
    - tpl 网页模板目录
    - uploads 上传文件目录
        - doc 最初保存docx文档的目录
        - html 最后保存转换出html的目录
        - image html文件里图片的目录
        - unzip 临时存放解压文件夹的目录
        - zip   临时存放转为zip包的目录

    - doc中附上三个样本例子，html/image中为对应转出来的页面和资源。

- /src/Docx2Html.php为主类文件，/srcexample.php为使用例子。

- 环境依赖：php需要载入zip/dom两个扩展，目前高版本php内核已经自动包含，低版本在php.ini中找到zip/dom的so(或dll)，去掉注释。

[![LICENSE](https://img.shields.io/badge/license-NPL%20(The%20996%20Prohibited%20License)-blue.svg)](https://github.com/996icu/996.ICU/blob/master/LICENSE)
[![996.icu](https://img.shields.io/badge/link-996.icu-red.svg)](https://996.icu)
    
