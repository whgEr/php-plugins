<?php  
/**
 * @author whgEr
 * OpenOffice 转office为pdf
 * 注意：仅在windows下测试成功(已测试:doc,xls,ppt,xlsx)
 * 下载地址：(widows) http://jaist.dl.sourceforge.net/project/openofficeorg.mirror/localized/zh-CN/3.4.1/Apache_OpenOffice_incubating_3.4.1_Win_x86_install_zh-CN.exe
 * 先开启php.ini中com组件
 * 安装openoffice后在命令行中运行如下命令：
 * 安装目录：soffice -headless -accept="socket,host=127.0.0.1,port=8100;urp;" -nofirststartwizard
 */
set_time_limit(0); 
function MakePropertyValue($name, $value, $osm) {
    $oStruct = $osm->Bridge_GetStruct("com.sun.star.beans.PropertyValue");
    $oStruct->Name = $name;
    $oStruct->Value = $value;
    return $oStruct;
} 
function word2pdf($doc_url, $output_url) {
    $osm = new COM("com.sun.star.ServiceManager") or die("Please be sure that OpenOffice.org is installed.n");
    $args = array(MakePropertyValue("Hidden", true, $osm));
    $oDesktop = $osm->createInstance("com.sun.star.frame.Desktop");
    $oWriterDoc = $oDesktop->loadComponentFromURL($doc_url, "_blank", 0, $args);
    $export_args = array(MakePropertyValue("FilterName", "writer_pdf_Export", $osm));
    $oWriterDoc->storeToURL($output_url, $export_args);
    $oWriterDoc->close(true);
} 
function getPdfPath() {
    $dir = "file:///E:/work/test/dompdf/";
		//Linux路径形式：file:/root/test_presentation.ppt   未测试
    $file = $dir. '11.doc';//源文件
    $output_file = $dir . '11.pdf';//目标文件
    if(file_exists($file)) {
			word2pdf($file, $output_file);
			if(file_exists($output_file))
				echo 'success';
			else 
				echo 'failed';
		}else 
				echo 'file not exists';
}
/*执行转换*/
getPdfPath();
?>
