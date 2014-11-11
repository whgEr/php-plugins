<?php  
/**
 * @author whgEr
 * pdf 转swf，windows下已测试
 * 安装SWFTools，下载地址：http://www.swftools.org/swftools-2013-04-09-1007.exe
 * 未安装中文相关语言包,未发现中文丢失现象
 * 另：写入批文件也可执行
 * ==========================
 	 + $handle = fopen($bat,'w') or die("can't open file");
	 + fwrite($handle,$command);
	 + fclose($handle);
	 + $oExec=exec($command);
	 + unlink($bat);
 * ==========================
 * 注意：$command 命令中的路径如包含空格，需加引号
 */
 
	$dir = "E:/work/test/dompdf/";//文件路径
	$pdf = $dir . '11.pdf';//源文件
	$swf = $dir . '11.swf';//目标文件
	$command = "\"D:/Program Files/SWFTools/pdf2swf.exe\"  -t " . $pdf . " -o  " . $swf ; //-s flashversion=9(未知作用)
  $shell = new COM("WSCript.Shell") ;
	$oExec = $shell->Run("cmd /C " . $command, 0, true);
 ?>
