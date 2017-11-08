<?php

/** 
 * 数据导出 
 * @param  array $title   标题行名称 
 * @param  array $data   导出数据
 * @param  string $fileName 文件名 
 * @param  string $savePath 保存路径 
 * @param  $type   是否下载  false--保存   true--下载 
 * @return string   返回文件全路径 
 * @throws PHPExcel_Exception 
 * @throws PHPExcel_Reader_Exception 
 */  
function exportExcel($title=array(), $data=array(), $fileName='', $savePath='./', $isDown=false){  
    include('PHPExcel.php');  // 加载excel类文件
    $obj = new PHPExcel();  // 获取excel类
  
    //横向单元格标识  
    $cellName = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ');  
      
    $obj->getActiveSheet(0)->setTitle('sheet');   //设置sheet名称  
    $_row = 1;   //设置纵向单元格标识  
    if($title){  
        $_cnt = count($title);  
        $obj->getActiveSheet(0)->mergeCells('A'.$_row.':'.$cellName[$_cnt-1].$_row);   //合并单元格  
        $obj->setActiveSheetIndex(0)->setCellValue('A'.$_row, '数据导出：'.date('Y-m-d H:i:s'));  //设置合并后的单元格内容  
        $_row++;  
        $i = 0;  
        foreach($title AS $v){   //设置列标题  
            $obj->setActiveSheetIndex(0)->setCellValue($cellName[$i].$_row, $v);  
            $i++;
        }  
        $_row++;
    }  
  
    //填写数据  
    if($data){  
        $i = 0;  
        foreach($data AS $_v){  
            $j = 0;  
            foreach($_v AS $_cell){  
                $obj->getActiveSheet(0)->setCellValue($cellName[$j] . ($i+$_row), $_cell);  
                $j++;  
            }  
            $i++;  
        }  
    }  
      
    //文件名处理  
    if(!$fileName){  
        $fileName = uniqid(time(),true);  
    }  
  
    $objWrite = PHPExcel_IOFactory::createWriter($obj, 'Excel2007');  
  
    if($isDown){   //网页下载  
        header('pragma:public');  
        header("Content-Disposition:attachment;filename=$fileName.xls");  
        $objWrite->save('php://output');exit;  
    }  
  
    $_fileName = iconv("utf-8", "gb2312", $fileName);   //转码  
    $_savePath = $savePath.$_fileName.'.xlsx';  
     $objWrite->save($_savePath);  
  
     return $savePath.$fileName.'.xlsx';  
}  

// 获取数据
$user = new PDO("mysql:host=localhost;dbname=excel","root","root");
$sql = "SELECT * FROM user";
$res = $user->query($sql)->fetchAll(PDO::FETCH_ASSOC);


// 将数据拆分成 key 和 val 两个数组
for ($i=0; $i < count($res); $i++) { 
    $key = array_keys($res[$i]);
    $val[] = array_values($res[$i]);
}


exportExcel($key, $val, '用户信息表', 'C:\Users\Public\Downloads', true);