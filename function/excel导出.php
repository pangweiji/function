<?php
//导出excel
$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
	->setLastModifiedBy("Maarten Balliauw")
	->setTitle("Office 2007 XLSX Test Document")
	->setSubject("Office 2007 XLSX Test Document")
	->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
	->setKeywords("office 2007 openxml php")
	->setCategory("Test result file");

$objPHPExcel->setActiveSheetIndex(0)->setCellValue('A1', '模板名称');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('B1', '内容');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('C1', '所属人');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('D1', '是否公用');
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(40);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(55);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10);

$index      = 2;    
foreach ($tplDataArr as $val){
	$objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$index, $val['name']);
	$objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$index, $val['content']);
	$objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$index, $val['username']);
	if ($val['iscommon'] == 1) {
		$iscommon = '是';
	} else {
		$iscommon = '否';
	}
	$objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$index, $iscommon);
	$index++; 
}

ob_end_clean();//清除缓冲区,避免乱码
header('Content-Type: application/vnd.ms-excel');  
header('Content-Disposition: attachment;filename="wish_template' . date('Y-m-d',time()) . '.xls"');  
header('Cache-Control: max-age=0'); 
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output'); 