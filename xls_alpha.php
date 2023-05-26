<?php
/*
 * Hello, I made a couple of methods that help me fill out the document using a loop that goes through a query to the database, 
 * this saves me work and it is less confusing for me to place the letters in the header, 
 * I don't know where to put the code, maybe you can organize it, but I hope this is useful as it is for me
 */

public function NotNull($array){
		#remove nulled or empty values
		foreach ($array as $key => $value) {
			if (is_array($value)) {
				$array[$key] = NotNull($array[$key]);
			}
			if (is_null($array[$key])) {
				unset($array[$key]);
			}
		}
		return $array;
	}

	public function xls_alpha($limit,$first_row=false){
		/*
		 * This helps me fill in the letters in Excel so I don't do it manually
		 * TEST: print_r(xls_alpha(50)); echo count(xls_alpha());
		 */
		$alpharray = [];
		$index=1;
		$loop = 0;
		$repeat = range('A','Z');

		while($index<=$limit){
			foreach($repeat as $v){
				#$alpharray[$index] = (($index<=$limit)? (($loop>1)? $v.($loop) : $v.($loop+1) ) :null);
				$alpharray[$index] =
					(($index<=$limit)?
						(
							($index<27)?
								$v
							:
								(
									($index%27==0)?	#After the first letter cycle A-Z, empiezan los multiplos
										$v.$v #.$repeat[1]
									:
										$repeat[$loop-1].$v
								)
						).(($first_row==true)?1:'')
							#When the second parameter is false, it is apt to fill the body of the document
							#if true, add a 1 to be the header
					:
						null	#No more letters, when exceeding the limit given by parameter
					);
				$index++;
			}
				$loop++;
		}
		#remove nulled or empty values
		#return ($alpharray = array_filter($alpharray, fn ($value) => !is_null($value)));	#only for PHP version >= 7.4 (Requires arrow function)
		return self::NotNull($alpharray);	#only for PHP version >= 5.6
		#return print_r($alpharray);
		
		
		#How To Use (Using While with MySQL data) : 
		/*
		
		
			#Step 1. Style

			$style_head = array(
				'font' => array(
					'bold' => true,
					'color' => array(
							'argb' => 'FFFFFF',
						),
				),
				'alignment' => array(
					'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,

				),
				'borders' => array(
					'top' => array(
						'style' => PHPExcel_Style_Border::BORDER_THIN,
					),
				),
				'fill' => array(
					'type' => PHPExcel_Style_Fill::FILL_SOLID,
					'color' => array(
						'argb' => '1D5091',
					),
				),
			);

			$style_body=array(
				'fill' => array(
				'type' => PHPExcel_Style_Fill::FILL_SOLID,
				'color' => array(
					'argb' => '9DDCFF',
					),
				),
			);


			#Step 2. Fill data

			$xls_data = [];
			$fetch = $db->query("SELECT * FROM empoyee");
			while ($row=$fetch->fetch_array()){
				$xls_data[] = array(
					'#ID'=> $row["id"],
					'Fisrt Name'=> $row["first_name"],
					'Last Name'=> $row["last_name"],
					'Email'=> $row["email"]
				);
			}

			#Step 3. Join Header with body
			if(is_array($xls_data) && isset($xls_data)){
				$test = '';	# [debug]
				$total_rows = count($xls_data[0]);	#The [0] reference the first sub array to count how many titles it has
				$alpha = $objPHPExcel->xls_alpha($total_rows, true); #for header
				$alpha2 = $objPHPExcel->xls_alpha($total_rows);	#for body

				$i = 1; #First row 1 (header)
				$col = 1; #Start looping through the column names of each row
				$last_col = '';	#Last column to add styles

				foreach($xls_data as $content){
					foreach($content as $h=>$v){
						$objPHPExcel->setActiveSheetIndex(0)->setCellValue("$alpha[$col]", "$h");
						$objPHPExcel->getActiveSheet()->getColumnDimension("$alpha[$col]")->setAutoSize(true);
						$last_col = $alpha[$col];
						$col++;	#loop through the column names of the current row (1=cabecera), then, $i jumps to the next row
					}	#end of header
					break;
				}

				$objPHPExcel->getActiveSheet()->getStyle("A1:$last_col")->applyFromArray($style_head);	#Head style
				$objPHPExcel->getActiveSheet()->setAutoFilter("A1:$last_col");	#Filters in the header

				$i = 2;	#From here, follow the body
				$row_style = 0;
				foreach($xls_data as $content){
					$col = 1; #Reset, to loop through each column header in the next loop that loops through the content
					foreach($content as $h=>$v){
						$objPHPExcel->setActiveSheetIndex(0)->setCellValue("$alpha2[$col]".$i, "$v");
						$col++;	#loop through the column names of the current row before $i jumps to the next row
					}
					$i++;	#next row
					#body style [test] Yes it works, but it applies to all columns without being limited to just the header
					
					if($row_style==1){
						$row_style=0; #Turn OFF
					} else {
						$objPHPExcel->getActiveSheet()->getStyle("$alpha2[$col]".$i)->applyFromArray($style_body);
						$row_style=1;	#Turn ON
					}
				}
			}
			
			#Step 4. Export
			...
	
		*/
		
		
	}
