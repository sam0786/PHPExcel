# PHPExcel
phpexcel yii2, php 7, xlsx


Descargar el paquete 

	composer  require --ignore-platform-reqs --prefer-dist sam0786/phpexcel:dev-master

ejemplo de uso en el controlador

	public function actionTestexcel(){
		
		$objPHPExcel = new \PHPExcel();
		
		$objSheet = $objPHPExcel->setActiveSheetIndex(0);
		$objPHPExcel->getProperties()
					->setCreator("reporte")
					->setLastModifiedBy("reporte")
					->setTitle("Reporte")
					->setSubject("Reporte")
					->setDescription("Reporte")
					->setCategory("Reportes");

		$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B1:Q1');
		$objPHPExcel->getActiveSheet()->getStyle('B1')->getAlignment()->applyFromArray(
			array('horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,)
		);

		$objPHPExcel->getActiveSheet()->getStyle("B1")->getFont()->setBold(true);

		$objSheet->setCellValue('B1', 'Titulo');
				//c4d79b
		$objPHPExcel->getActiveSheet()
					->getStyle('B2:Q2')
					->applyFromArray(
						array(
							'fill' => array(
								'type' => \PHPExcel_Style_Fill::FILL_SOLID,
								'color' => array('rgb' => '060606')
							)
						)
					);

		$styleArray = array(
				  'borders' => array(
					'allborders' => array(
						'style' => \PHPExcel_Style_Border::BORDER_THIN,
						'color' => array('rgb' => '000000')
					)
				  )
		);
		
		$styleArrayt = array(
				'font'  => array(
					'bold'  => true,
					'color' => array('rgb' => 'ffffff'),
					'size'  => 12,
		));
		
		$objPHPExcel->getActiveSheet()->getStyle("B2:Q2")->applyFromArray($styleArray);
		$objPHPExcel->getActiveSheet()->getStyle("B2:Q2")->applyFromArray($styleArrayt);
			
				
		
				$objPHPExcel->setActiveSheetIndex(0);
				$objPHPExcel->getActiveSheet()->SetCellValue('B2', 'TEXT 1');
				$objPHPExcel->getActiveSheet()->SetCellValue('C2', 'TEXT 2');
				$objPHPExcel->getActiveSheet()->SetCellValue('D2', 'TEXT 3');
				$objPHPExcel->getActiveSheet()->SetCellValue('E2', 'TEXT 4');
				$objPHPExcel->getActiveSheet()->SetCellValue('F2', 'TEXT 5');
				$objPHPExcel->getActiveSheet()->SetCellValue('G2', 'TEXT 6');
				$objPHPExcel->getActiveSheet()->SetCellValue('H2', 'TEXT 7');
				$objPHPExcel->getActiveSheet()->SetCellValue('I2', 'TEXT 8');
				$objPHPExcel->getActiveSheet()->SetCellValue('J2', 'TEXT 9');
				$objPHPExcel->getActiveSheet()->SetCellValue('K2', 'TEXT 10');
				$objPHPExcel->getActiveSheet()->SetCellValue('L2', 'TEXT 11');
				$objPHPExcel->getActiveSheet()->SetCellValue('M2', 'TEXT 12');
				$objPHPExcel->getActiveSheet()->SetCellValue('N2', 'TEXT 13');
				$objPHPExcel->getActiveSheet()->SetCellValue('O2', 'TEXT 14');
				$objPHPExcel->getActiveSheet()->SetCellValue('P2', 'TEXT 15');
				$objPHPExcel->getActiveSheet()->SetCellValue('Q2', 'TEXT 16');
		
				

				$objPHPExcel->getActiveSheet()->setTitle('Reporte SGI');
				$objPHPExcel->setActiveSheetIndex(0);

			Yii::$app->response->headers->add('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
			header('Content-Disposition: attachment;filename="reporte.xlsx"');
			header('Cache-Control: max-age=0');
			$objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
			$objWriter->save('php://output');			
		 	exit;
	}
