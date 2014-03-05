<?php

error_reporting(E_ALL);
ini_set('display_errors', true);
ini_set('display_startup_errors', true);
set_time_limit(0);
date_default_timezone_set('Europe/Moscow');

define('EOL', (PHP_SAPI == 'cli') ? PHP_EOL : '<br/>');

require_once 'func.php';
?>

<!DOCTYPE html>
<html>
	<head></head>
	<body>
		<form action="" enctype="multipart/form-data" method="POST">
			<table>
				<tr>
					<td>
						Выберите файл с шаблоном
					</td>
					<td>
						<input type="file" name="template_file" />
					</td>
				<tr>
					<td>
						Выберите файлы с афишей
					</td>
					<td>
						<input type="file" name="file_xls[]" multiple="true" />
					</td>
				</tr>
				<tr>
					<td>
						<input type="submit" value="Запустить обработку" />
					</td>
					<td></td>
				</tr>
				</tr>
			</table>
			<input type="hidden" name="MAX_FILE_SIZE" value="50000">
		</form>
	</body>
</html>

<?php
	if (isset($_FILES['template_file']['tmp_name']) && isset($_FILES['file_xls']['tmp_name']))
	{
		echo 'Файл с шаблоном: ' . $_FILES['template_file']['name'] . '<br/>';
		$fileArray = arrayFiles($_FILES['file_xls']);
		foreach ($fileArray as $file)
		{
			echo 'File Name: ' . $file['name'] . '<br/>';
			echo 'File tempname: ' . $file['tmp_name'] . '<br/>';
		}
		$uploadDir = __DIR__ . '/';
		$uploadFile = $uploadDir . basename($_FILES['template_file']['name']);
		if (!move_uploaded_file($_FILES['template_file']['tmp_name'], $uploadFile))
			echo 'Невозможно загрузить файл!';
		
		foreach ($fileArray as $file)
		{
			$uploadFile = $uploadDir . basename($file['name']);
			if (!move_uploaded_file($file['tmp_name'], $uploadFile))
				echo 'Невозможно загрузить файл!';
		}
//		require_once 'Classes/PHPExcel.php';
//		require_once 'Classes/PHPExcel/Writer/Excel2007.php';
//		require_once 'Classes/PHPExcel/IOFactory.php';
//		
//		$templateFilename = $_FILES['template_file']['tmp_name'];
//		$templateXls = PHPExcel_IOFactory::load($templateFilename);
//		$templateXls->getActiveSheetIndex(1);
//		$templateXlsSheet = $templateXls->getActiveSheet()->toArray();
//		echo '<pre>';
//		print_r($templateXlsSheet);
//		echo '</pre>';
	}
?>
<?php
//
//$sheet = array();
//
//$fileName = __DIR__ . '/SinemaStar.xls';
//
//if (empty($fileName))
//	throw new Exception("No file specified.");
//
//if (!file_exists($fileName))
//	throw new Exception("Could not open " . $fileName . " for reading! File does not exist.");
//
//$xls = PHPExcel_IOFactory::load($fileName);
//
//$xls->setActiveSheetIndex(0);
//
//$sheet = $xls->getActiveSheet()->toArray();
//
//$xls = null;
//
//$halls = array();
//$hall = '';
//
//$timeTable = $sheet[0][0];
//$timeTable = normalizeTimeTable($timeTable);
//
//foreach ($sheet as $key => $value)
//{
//	
//	if (strstr((string)$value[1], 'Зал:'))
//	{
//		$tmp = explode(',', $value[1]);
//		$hall = $tmp[0];
//	}
//	
//	if (strstr((string)$value[1], 'Сеанс') || strstr((string)$value[1], 'Зал:') || $value[1] === null)
//	{
//		unset($sheet[$key]);
//		continue;
//	}
//	
//	if (!$hall == '')
//		$halls[$hall][] = $value;
//}
//
//$sheet = array_values($sheet);
//
//$hallFilms = array();
//foreach ($halls as $roomHall => $hall)
//{
//	for ($i=0; $i<count($hall); $i++)
//	{
//		if ($hall[$i][3] == null)
//			continue;
//		
//		$filmName = $hall[$i][3];
//		$hallFilms[$roomHall][$filmName] = '';
//		
//		for ($j=0;$j<count($hall);$j++)
//		{
//			if ($hall[$j][3] === $filmName)
//			{
//				$hallFilms[$roomHall][$filmName] .= $hall[$j][1] . ',';
//				$hall[$j][3] = null;
//			}
//		}
//	}
//}
//
//$xls = PHPExcel_IOFactory::load(__DIR__ . '/афиша-заливка.xls');
//$xls->setActiveSheetIndex(0);
//
//$startWith = 3;
//
//foreach ($hallFilms as $hallName => $films)
//{
//	$counter = 0;
//	
//	foreach ($films as $filmName => $filmTime)
//	{
//		$cellId = $startWith + $counter;
//		
//		$xls->getActiveSheet()
//			->setCellValue('A'.$cellId, normalizeFilmName($filmName))
//			->setCellValue('B'.$cellId, $timeTable['startDate'])
//			->setCellValue('C'.$cellId, $timeTable['endDate'])
//			->setCellValue('D'.$cellId, normalizeTime($filmTime))
//			->setCellValue('F'.$cellId, normalizeHallName($hallName));
//		
//		$counter++;
//	}
//	
//	$startWith = $cellId + 1;
//}
//
//$objWriter = PHPExcel_IOFactory::createWriter($xls, 'Excel5');
//$objWriter->save(__DIR__ . '/афиша-заливка.xls');