<?php

error_reporting(E_ALL);
ini_set('display_errors', true);
ini_set('display_startup_errors', true);
set_time_limit(0);
date_default_timezone_set('Europe/Moscow');

define('EOL', (PHP_SAPI == 'cli') ? PHP_EOL : '<br/>');

function normalizeTime($time)
{
	$timeArray = explode(',', $time);
	
	foreach ($timeArray as $key => $value)
	{
		if ($value == '')
			unset($timeArray[$key]);
	}
	
	return implode(',', $timeArray);
}

function normalizeFilmName($filmName)
{
	$excess = array(
		'*',
		'#',
		'2D',
		'3D',
		'5D',
		'7D',
		'12+',
		'14+',
		'16+',
		'18+',
		'0+',
		'6+',
		'1',
		'2',
	);
	
	$filmName = str_ireplace($excess, '', $filmName);
	
	return trim($filmName);
}

function normalizeHallName($hallName)
{
	return str_replace(':', '', $hallName);
}

function normalizeTimeTable($timeTable, $type=null)
{
	if ($type == null)
		return $timeTable;
	elseif ($type == 'sinemastar') 
	{
		$timeTable = filter_var($timeTable, FILTER_SANITIZE_NUMBER_INT);
		$timeTable = str_replace('-', '', $timeTable);

		$year = substr($timeTable, -4, 4);
		$month = substr($timeTable, -6, 2);
		$dayOn = substr($timeTable, -8, 2);
		$dayWith = substr($timeTable, 0, 2);

		return array(
			'startDate' => $dayWith . '.' . $month . '.' . $year,
			'endDate' => $dayOn . '.' . $month . '.' . $year,
		);	
	}
	elseif ($type == 'sinemapark')
	{
		$year = '20' . substr($timeTable, 6, 2);
		$month = substr($timeTable, 3, 2);
		$dayOn = substr($timeTable, 0, 2);
		$dayWith = substr($timeTable, 0, 2);
		
		return array(
			'startDate' => $dayWith . '.' . $month . '.' . $year,
			'endDate' => $dayOn . '.' . $month . '.' . $year,
		);	
	}
	elseif ($type == 'karorussia')
	{
		$date = explode('-', $timeTable);
		
		return array(
			'startDate' => $date[0],
			'endDate' => $date[1],
		);		
	}
}

function arrayFiles(&$filePost) 
{
    $fileArray = array();
    $fileCount = count($filePost['name']);
    $fileKeys = array_keys($filePost);

    for ($i=0; $i < $fileCount; $i++) 
	{
        foreach ($fileKeys as $key) 
		{
            $fileArray[$i][$key] = $filePost[$key][$i];
        }
    }

    return $fileArray;
}

?>

<!DOCTYPE html>
<html>
	<head></head>
	<body>
		<form action="" enctype="multipart/form-data" method="POST">
			<table>
				<tr>
					<td>
						Сгенерировать шаблон?
					</td>
					<td>
						<input type="submit" name="generate_template" value="Да" />
					</td>
				</tr>
				<tr>
					<td>
						Выберите файл с шаблоном
					</td>
					<td>
						<input type="file" name="template_file" />
					</td>
				</tr>
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
	if (isset($_POST['generate_template']))
	{
		require_once 'Classes/PHPExcel.php';
		require_once 'Classes/PHPExcel/Writer/Excel2007.php';

		$objPHPExcel = new PHPExcel();
		$objPHPExcel->setActiveSheetIndex(0);
		$objPHPExcel->getActiveSheet()->setTitle('кино_заливка');
		$objPHPExcel->getActiveSheet()
					->setCellValue('A1', '#заголовок')
					->setCellValue('B1', '#Дата начала афиша')
					->setCellValue('C1', '#Дата окончания афиша')
					->setCellValue('D1', '#Сеансы')
					->setCellValue('E1', '#Привязка описания события')
					->setCellValue('F1', '#Зал_new');
		$objPHPExcel->getActiveSheet()
					->setCellValue('A2', 'mess_header')
					->setCellValue('B2', 'comm_date_min_new')
					->setCellValue('C2', 'comm_date_max_new')
					->setCellValue('D2', 'comm_session')
					->setCellValue('E2', 'af_repert_mn')
					->setCellValue('F2', 'af_hall_new');

		$objPHPExcel->createSheet(1);
		$objPHPExcel->setActiveSheetIndex(1);
		$objPHPExcel->getActiveSheet()->setTitle('события');

		$objPHPExcel->createSheet(2);
		$objPHPExcel->setActiveSheetIndex(2);
		$objPHPExcel->getActiveSheet()->setTitle('залы');
		$objPHPExcel->getActiveSheet()
					->setCellValue('A1', '---')
					->setCellValue('A2', 'Зал 1')
					->setCellValue('A3', 'Зал 2')
					->setCellValue('A4', 'Зал 3')
					->setCellValue('A5', 'Зал 4')
					->setCellValue('A6', 'Зал 5')
					->setCellValue('A7', 'Зал 6')
					->setCellValue('A8', 'Зал 7')
					->setCellValue('A9', 'Зал 8')
					->setCellValue('A10', 'Зал 9')
					->setCellValue('A11', 'Зал 10')
					->setCellValue('A12', 'Зал 11')
					->setCellValue('A13', 'Зал 12')
					->setCellValue('A14', 'Зал 13')
					->setCellValue('A15', 'Зал 14')
					->setCellValue('A16', 'Зал 15')
					->setCellValue('A17', 'Зал 16')
					->setCellValue('A18', 'Зал 17')
					->setCellValue('A19', 'Зал 18')
					->setCellValue('A20', 'Зал 19')
					->setCellValue('A21', 'Зал 20')
					->setCellValue('A22', 'Большой зал')
					->setCellValue('A23', 'Малый зал')
					->setCellValue('A24', 'Синий зал')
					->setCellValue('A25', 'Зеленый зал')
					->setCellValue('A26', 'Vip зал')
					->setCellValue('A27', 'Большой звездный')
					->setCellValue('A28', 'Космонавтика')
					->setCellValue('A29', 'Астрономия')
					->setCellValue('A30', 'dk')
					->setCellValue('A31', 'da')
					->setCellValue('A32', 'Зал Relax')
					->setCellValue('A33', 'Зал IMAX')
					->setCellValue('A34', 'Зал Jolly')
					->setCellValue('A35', 'Зал 4DX')
					->setCellValue('A36', 'Кремлевский концертный зал')
					->setCellValue('A37', 'Концертный зал Консерватории');

		$objPHPExcel->getActiveSheet()
					->setCellValue('B1', '---')
					->setCellValue('B2', 'Зал 1')
					->setCellValue('B3', 'Зал 2')
					->setCellValue('B4', 'Зал 3')
					->setCellValue('B5', 'Зал 4')
					->setCellValue('B6', 'Зал 5')
					->setCellValue('B7', 'Зал 6')
					->setCellValue('B8', 'Зал 7')
					->setCellValue('B9', 'Зал 8')
					->setCellValue('B10', 'Зал 9')
					->setCellValue('B11', 'Зал 10')
					->setCellValue('B12', 'Зал 11')
					->setCellValue('B13', 'Зал 12')
					->setCellValue('B14', 'Зал 13')
					->setCellValue('B15', 'Зал 14')
					->setCellValue('B16', 'Зал 15')
					->setCellValue('B17', 'Зал 16')
					->setCellValue('B18', 'Зал 17')
					->setCellValue('B19', 'Зал 18')
					->setCellValue('B20', 'Зал 19')
					->setCellValue('B21', 'Зал 20')
					->setCellValue('B22', 'Большой зал')
					->setCellValue('B23', 'Малый зал')
					->setCellValue('B24', 'Синий зал')
					->setCellValue('B25', 'Зеленый зал')
					->setCellValue('B26', 'Vip зал')
					->setCellValue('B27', 'Большой звездный')
					->setCellValue('B28', 'Космонавтика')
					->setCellValue('B29', 'Астрономия')
					->setCellValue('B30', 'ДК ОАО "Завод "Красное Сормово"')
					->setCellValue('B31', 'Дом Актера')
					->setCellValue('B32', 'Зал Relax')
					->setCellValue('B33', 'Зал IMAX')
					->setCellValue('B34', 'Зал Jolly')
					->setCellValue('B35', 'Зал 4DX')
					->setCellValue('B36', 'Кремлевский концертный зал')
					->setCellValue('B37', 'Концертный зал Консерватории');

		$objPHPExcel->setActiveSheetIndex(0);
		
		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="афиша-заливка.xls"');
		header('Cache-Control: max-age=0');
		header('Cache-Control: max-age=1');
		header('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT');
		header('Cache-Control: cache, must-relavidate');
		header('Pragma: public');
		
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
		ob_clean();
		$objWriter->save('php://output');
		exit;
	}

	if (isset($_FILES['template_file']['tmp_name']) && isset($_FILES['file_xls']['tmp_name']))
	{
		$uploadDir = __DIR__ . '/';
		$uploadFile = $uploadDir . basename($_FILES['template_file']['name']);
		
		if (!move_uploaded_file($_FILES['template_file']['tmp_name'], $uploadFile))
			echo 'Невозможно загрузить файл!';
		
		$fileArray = arrayFiles($_FILES['file_xls']);
		
		foreach ($fileArray as $file)
		{
			$uploadFile = $uploadDir . basename($file['name']);
			if (!move_uploaded_file($file['tmp_name'], $uploadFile))
				echo 'Невозможно загрузить файл!';
		}
		
		require_once 'Classes/PHPExcel.php';
		require_once 'Classes/PHPExcel/Writer/Excel2007.php';
		require_once 'Classes/PHPExcel/IOFactory.php';
		
		$templateFilename = $uploadDir . 'афиша-заливка.xls';
		$templateXls = PHPExcel_IOFactory::load($templateFilename);
		$templateXls->setActiveSheetIndex(1);
		$templateXls->getActiveSheet();
		$eventsSheet = $templateXls->getActiveSheet()->toArray();
		$templateXls->setActiveSheetIndex(0);
		$templateXls->getActiveSheet();
		$templateXlsSheet = $templateXls->getActiveSheet()->toArray();
		
		$startWith = 3;
		
		foreach ($fileArray as $file)
		{
			$fileName = __DIR__ . '/' . $file['name'];
			
			switch (pathinfo($fileName, PATHINFO_FILENAME))
			{
				case 'SinemaStar':
					if (!file_exists($fileName))
						throw new Exception("Could not open " . $fileName . " for reading! File does not exist.");
					
					$afisha = PHPExcel_IOFactory::load($fileName);
					$afisha->setActiveSheetIndex(0);
					$afishaSheet = $afisha->getActiveSheet()->toArray();
					$afisha = null;

					$halls = array();
					$hall = '';

					$timeTable = normalizeTimeTable($afishaSheet[0][0], 'sinemastar');

					foreach ($afishaSheet as $key => $value)
					{

						if (strstr((string)$value[1], 'Зал:'))
						{
							$tmp = explode(',', $value[1]);
							$hall = $tmp[0];
						}

						if (strstr((string)$value[1], 'Сеанс') || strstr((string)$value[1], 'Зал:') || $value[1] === null)
						{
							unset($afishaSheet[$key]);
							continue;
						}

						if (!$hall == '')
							$halls[$hall][] = $value;
					}
					
					$hallFilms = array();
					
					foreach ($halls as $roomHall => $hall)
					{
						for ($i=0; $i<count($hall); $i++)
						{
							if ($hall[$i][3] == null)
								continue;

							$filmName = $hall[$i][3];
							$hallFilms[$roomHall][$filmName] = '';

							for ($j=0;$j<count($hall);$j++)
							{
								if ($hall[$j][3] === $filmName)
								{
									$hallFilms[$roomHall][$filmName] .= $hall[$j][1] . ',';
									$hall[$j][3] = null;
								}
							}
						}
					}
					
					foreach ($hallFilms as $hallName => $films)
					{
						$counter = 0;

						foreach ($films as $filmName => $filmTime)
						{
							$cellId = $startWith + $counter;
							
							foreach ($eventsSheet as $event)
							{
								if (trim($event[1]) == normalizeFilmName($filmName))
								{
									$eventId = $event[0];
									break;
								}
								else
									$eventId = 'Событие не найдено!';
							}

							$templateXls->getActiveSheet()
										->setCellValue('A'.$cellId, normalizeFilmName($filmName))
										->setCellValue('B'.$cellId, $timeTable['startDate'])
										->setCellValue('C'.$cellId, $timeTable['endDate'])
										->setCellValue('D'.$cellId, normalizeTime($filmTime))
										->setCellValue('E'.$cellId, $eventId)
										->setCellValue('F'.$cellId, normalizeHallName($hallName));

							$counter++;
						}

						$startWith = $cellId + 1;
					}
					
					$objWriter = PHPExcel_IOFactory::createWriter($templateXls, 'Excel5');
					$objWriter->save(__DIR__ . '/sinema_star_mod.xls');
					$objPHPExcel = null;
					exit;
					break;
				case 'SinemaPark':
					if (!file_exists($fileName))
						throw new Exception("Could not open " . $fileName . " for reading! File does not exist.");
					
					$afisha = PHPExcel_IOFactory::load($fileName);
					$afisha->setActiveSheetIndex(0);
					$afishaSheet = $afisha->getActiveSheet()->toArray();
					$afisha = null;
					
					$timeTable = normalizeTimeTable($afishaSheet[2][0], 'sinemapark');
					
					$halls = array();
					$hall = '';
					
					foreach ($afishaSheet as $key => $value)
					{
						if (strstr($value[0], 'Зал'))
							$hall = $value[0];

						if (($value[1] == '' && $value[2] == '') || strstr($value[0], 'Зал') || strstr($value[1], '*'))
						{
							unset($afishaSheet[$key]);
							continue;
						}

						if (!$hall == '')
							$halls[$hall][] = $value;
					}
					
					foreach ($halls as $roomHall => $hall)
					{
						$counter = 0;
						
						foreach ($hall as $film)
						{
							$cellId = $startWith + $counter;
							$filmTime = '';
							
							for ($i=3; $i<29; $i++)
							{
								if ($film[$i] !== '')
									$filmTime .= $film[$i] . ',';
							}
							
							foreach ($eventsSheet as $event)
							{
								if (trim($event[1]) == normalizeFilmName($film[1]))
								{
									$eventId = $event[0];
									break;
								}
								else
									$eventId = 'Событие не найдено!';
							}
							
							$templateXls->getActiveSheet()
										->setCellValue('A'.$cellId, normalizeFilmName($film[1]))
										->setCellValue('B'.$cellId, $timeTable['startDate'])
										->setCellValue('C'.$cellId, $timeTable['endDate'])
										->setCellValue('D'.$cellId, normalizeTime($filmTime))
										->setCellValue('E'.$cellId, $eventId)
										->setCellValue('F'.$cellId, $roomHall);
							
							$counter++;
						}
						
						$startWith = $cellId + 1;
					}
					
					
					$objWriter = PHPExcel_IOFactory::createWriter($templateXls, 'Excel5');
					$objWriter->save(__DIR__ . '/sinema_park_mod.xls');
					$objPHPExcel = null;
					exit;
					break;
				case 'KaroRussia':
					if (!file_exists($fileName))
						throw new Exception("Could not open " . $fileName . " for reading! File does not exist.");
					
					$afisha = PHPExcel_IOFactory::load($fileName);
					$afishaSheet = $afisha->getActiveSheet()->toArray();
					$afisha = null;
					
					$timeTable = normalizeTimeTable($afishaSheet[2][0], 'karorussia');
					
					$halls = array();
					$hall = '';
					
					foreach ($afishaSheet as $key => $value)
					{
						if ($value[0] == '1' || $value[0] == '2' || $value[0] == '3')
						{
							$hall = 'Зал '.$value[0];
						}
						
						if ($value[2] != '')
							$halls[$hall][] = $value;
						
						unset($afishaSheet[$key]);
					}
					unset($halls['']);
					
					foreach ($halls as $roomHall => $hall)
					{
						$counter = 0;
						
						foreach ($hall as $film)
						{
							$cellId = $startWith + $counter;
							$filmTime = '';
							
							for ($i=3; $i<16; $i++)
							{
								if ($film[$i] !== '')
									$filmTime .= $film[$i] . ',';
							}
							
							if ($film[2] == 'Подготовлено:' || $film[2] == 'Изменено:')
							{
								unset($film);
								break;
							}
							
							foreach ($eventsSheet as $event)
							{
								if (trim($event[1]) == normalizeFilmName($film[2]))
								{
									$eventId = $event[0];
									break;
								}
								else
									$eventId = 'Событие не найдено!';
							}
							
							$templateXls->getActiveSheet()
										->setCellValue('A'.$cellId, normalizeFilmName($film[2]))
										->setCellValue('B'.$cellId, $timeTable['startDate'])
										->setCellValue('C'.$cellId, $timeTable['endDate'])
										->setCellValue('D'.$cellId, normalizeTime($filmTime))
										->setCellValue('E'.$cellId, $eventId)
										->setCellValue('F'.$cellId, $roomHall);
							
							$counter++;
						}
						
						$startWith = $cellId + 1;
					}
					
										
					$objWriter = PHPExcel_IOFactory::createWriter($templateXls, 'Excel5');
					$objWriter->save(__DIR__ . '/karo_russia_mod.xls');
					$objPHPExcel = null;
					break;
				default:
					break;
			}
		}
	}
?>