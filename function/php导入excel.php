<?php
$file = $_FILES['file'];
if($file['error'] == 0){
	$file_type = end(explode('.',$file['name']));
	$ext = array('xls','xlsx');
	
	//判断excel文件
	if (!in_array(strtolower($file_type),$ext)) {
		self::$errMsg['100'] = "该文件格式错误，请导入excel格式文件！";
		return false;
	}

	$save_path = WEB_PATH. 'html/upload/';
	if (!file_exists($save_path)) {
		mkdir($save_path);
	}
	$file_path = time().$company_id.'importSPU.'.$file_type;
	$filePath = $save_path.$file_path;
	if (!move_uploaded_file($file['tmp_name'], $filePath)) {
		self::$errMsg['100'] = "导入失败，文件读取错误！";
		return false;
	}

	//转化成数组
	$objPHPExcel = E('PHPExcel');
	$PHPReader = new PHPExcel_Reader_Excel2007();
	if (!$PHPReader->canRead($filePath)) {
		$PHPReader = new PHPExcel_Reader_Excel5();
		if (!$PHPReader->canRead($filePath)) {
			unlink($filePath);
			self::$errMsg['100'] = "导入失败，文件读取错误！";
			return false;
		}
	}
	$objPHPExcel    = $PHPReader->load($filePath);
	$currentSheet   = $objPHPExcel->getSheet(0);
	$allColumn      = $currentSheet->getHighestColumn();
	$allRow         = $currentSheet->getHighestRow();

	$data = array();
	$nowYear = date('Y', time());
	for ($currentRow = 2;$currentRow <= $allRow;$currentRow++) {
		for ($currentColumn = 'A';$currentColumn <= $allColumn;$currentColumn++) {
			$addr = $currentColumn.$currentRow;
			$cell = $currentSheet->getCell($addr)->getValue();
			if ($cell instanceof PHPExcel_RichText) {
				$cell = trim($cell->__toString());
			}
			if (empty($cell)) {
				unlink($filePath);
				self::$errMsg['100'] = '导入失败，' . "该文件{$currentColumn}列第{$currentRow}行数据不为空，请核查！";
				return false;
			}
			$formatFail = false;
			switch ($currentColumn) {
				case 'A':
					$data[$currentRow]['spu'] = $cell;
					break;
				case 'B':
					$time1 = PHPExcel_Shared_Date::ExcelToPHP($cell);
					if (!is_numeric($cell) || date('Y', $time1) > $nowYear+5 || date('Y', $time1) < $nowYear-2) {
						unlink($filePath);
						self::$errMsg['100'] = '导入失败，' . "该文件{$currentColumn}列第{$currentRow}行数据格式错误，请核查！";
						return false;
					}
					$data[$currentRow]['effecttime'] = $time1 - 28800;
					break;
				case 'C':
					$time2 = PHPExcel_Shared_Date::ExcelToPHP($cell);
					if (!is_numeric($cell) || date('Y', $time2) > $nowYear+5 || date('Y', $time2) < $nowYear-2) {
						unlink($filePath);
						self::$errMsg['100'] = '导入失败，' . "该文件{$currentColumn}列第{$currentRow}行数据格式错误，请核查！";
						return false;
					}
					$data[$currentRow]['failtime'] = $time2 - 28800 + 86399;
					break;
			}
		}
		if ($time1 > $time2) {
			unlink($filePath);
			self::$errMsg['100'] = '导入失败，该文件第'.$currentRow.'行生效时间不能大于失效时间！';
			return false;
		}
	}
	unlink($filePath);
}