<?php
/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2011 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2011 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.7.6, 2011-02-27
 */

/** Error reporting */
error_reporting(E_ALL);
//aca hize un cambio
date_default_timezone_set('Europe/London');

/** PHPExcel */
require_once 'models/vistaCurso.php';
require_once 'models/respuesta.php';
require_once 'models/facultad.php';
require_once 'models/escuela.php';
require_once 'Classes/PHPExcel.php';
$objPHPExcel = new PHPExcel();

// Set properties
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
                             ->setLastModifiedBy("Maarten Balliauw")
                             ->setTitle("Office 2007 XLSX Test Document")
                             ->setSubject("Office 2007 XLSX Test Document")
                             ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
                             ->setKeywords("office 2007 openxml php")
                             ->setCategory("Test result file");
// IGUAL Q EL CONTROLADOR CREPORTE =)
$annio= $_REQUEST['annio'];
$periodo= $_REQUEST['periodo'];
$ciclo= $_REQUEST['ciclo'];
$facu= $_REQUEST['facu'];
$esc= $_REQUEST['esc'];
	
    
    $cursos= new vistaCurso('', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '');
    $respuestas= new respuesta('', '','','','','','', '','');
    $facultad= new facultad('', '');
    $escuela= new escuela('', '','');
    $camposMostrar='DISTINCT(NOMCUR)*NOMDOC*APEDOC*YYCUR*CODCURR*CODCUR*CODDOCE*CODMODALIDAD*NOMMODALIDAD*YYCURR';
    $campo="YYAKD*CODPER*CODFAC*CODESC*YYCUR";
    $operador="=*=*=*=*=";
    $valor="$annio*$periodo*$facu*$esc*$ciclo";
    $separador="AND*AND*AND*AND";
   $c=$cursos->listarSimpleR($camposMostrar, $campo, $operador, $valor, $separador, '', '');
	$na=count($c);
    if($na>0){
            $fac = $facultad->listarSimple('NOMFAC', 'CODFAC', '=', $facu, '', '', '');
            $es = $escuela->listarSimple('NOMESC', 'CODFAC*CODESC', '=*=', $facu.'*'.$esc, 'AND', '', '');
	  /*
                             * WHERE asig.`CODESC`='04' 
                         * AND asig.CODFAC='01' 
                         * AND asig.YYAKD='2011' 
                         * AND asig.CODPER='2'
                         * AND asig.YYCUR='08' 
                           AND re.CODCURR='0096'
                         * AND re.CODCUR='03084'
                         * AND re.CODDOCE='FA0001' 
                         * AND pre.`dimensiones_id`='1'
                             */
                            $objPHPExcel->setActiveSheetIndex(0);
                            // Add a drawing to the worksheet
                            $objDrawing = new PHPExcel_Worksheet_Drawing();
                            $objDrawing->setName('Logo');
                            $objDrawing->setDescription('Logo');
                            $objDrawing->setPath('./images/logo2.png');
                            $objDrawing->setHeight(100);
                            $objDrawing->setCoordinates('A1');
                            $objDrawing->setWorksheet($objPHPExcel->getActiveSheet());


                            // Add rich-text string
                            $objRichText = new PHPExcel_RichText();
                            $objPayable = $objRichText->createTextRun('UNIVERSIDAD JOSÉ CARLOS MARIÁTEGUI');
                            $objPayable->getFont()->setBold(true);
                            $objPayable->getFont()->setName('Cooper Black');
                            $objPayable->getFont()->setSize(28);
                            $objPayable->getFont()->getColor()->setARGB("31869B");
                            $objPHPExcel->getActiveSheet()->getCell('B3')->setValue($objRichText);
                            $objPHPExcel->getActiveSheet()->mergeCells('B3:O4');

                            // Merge cells
                            $objPHPExcel->getActiveSheet()->getStyle('B3:O4')->applyFromArray(
                                        array(
                                            
                                            'alignment' => array(
                                                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                            )
                                        )
                                );

                                // 2do titulo
                            $objRichText = new PHPExcel_RichText();
                            $objPayable = $objRichText->createTextRun('OFICINA DE CALIDAD ACADÉMICA AUTOEVALUACIÓN Y ACREDITACIÓN UNIVERSITARIA');
                            $objPayable->getFont()->setBold(true);
                            $objPayable->getFont()->setSize(10);
                            $objPHPExcel->getActiveSheet()->getCell('C5')->setValue($objRichText);
                            $objPHPExcel->getActiveSheet()->mergeCells('C5:H5');

                            // Merge cells
                            $objPHPExcel->getActiveSheet()->getStyle('C5:H5')->applyFromArray(
                                        array(
                                            
                                            'alignment' => array(
                                                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                            )
                                        )
                                );

                                //3ero
                            $objPHPExcel->getActiveSheet()->setCellValue('C9', 'FACULTAD:');
                            $objPHPExcel->getActiveSheet()->getStyle('C9')->getFont()->setBold(true);
                            $objPHPExcel->getActiveSheet()->getStyle('C9')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                            $objRichText = new PHPExcel_RichText();
                            $objPayable = $objRichText->createTextRun($fac[0][0]);
                            $objPayable->getFont()->setBold(true);
                            $objPayable->getFont()->getColor()->setARGB("000000");
                            $objPHPExcel->getActiveSheet()->getCell('D9')->setValue($objRichText);
                            $objPHPExcel->getActiveSheet()->mergeCells('D9:J9');

                            // Merge cells
                            $objPHPExcel->getActiveSheet()->getStyle('D9:J9')->applyFromArray(
                                        array(
                                            
                                            'alignment' => array(
                                                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
                                            )
                                        )
                                );

                            $objPHPExcel->getActiveSheet()->setCellValue('C10', 'CARRERA:');
                            $objPHPExcel->getActiveSheet()->getStyle('C10')->getFont()->setBold(true);
                            $objPHPExcel->getActiveSheet()->getStyle('C10')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                            $objRichText = new PHPExcel_RichText();
                            $objPayable = $objRichText->createTextRun($es[0][0]);
                            $objPayable->getFont()->setBold(true);
                            $objPayable->getFont()->getColor()->setARGB("000000");
                            $objPHPExcel->getActiveSheet()->getCell('D10')->setValue($objRichText);
                            $objPHPExcel->getActiveSheet()->mergeCells('D10:J10');

                            // Merge cells
                            $objPHPExcel->getActiveSheet()->getStyle('D10:J10')->applyFromArray(
                                        array(
                                            
                                            'alignment' => array(
                                                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
                                            )
                                        )
                                );

                            $objPHPExcel->getActiveSheet()->setCellValue('K9', 'CICLO: '.$ciclo);
                             $objPHPExcel->getActiveSheet()->getStyle('K9')->getFont()->setBold(true);


////////////////////

                                $objRichText = new PHPExcel_RichText();
                                $objPayable = $objRichText->createTextRun('N°');
                                $objPayable->getFont()->setBold(true);
                                $objPHPExcel->getActiveSheet()->getCell('A13')->setValue($objRichText);
                                $objPHPExcel->getActiveSheet()->mergeCells('A13:A14');
                                $objPHPExcel->getActiveSheet()->getStyle('A13:A14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'name'      => 'Chaparral Pro',
                                            'size'      => '14'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objRichText = new PHPExcel_RichText();
                                $objPayable = $objRichText->createTextRun('Curso');
                                $objPayable->getFont()->setBold(true);
                                $objPHPExcel->getActiveSheet()->getCell('B13')->setValue($objRichText);
                                $objPHPExcel->getActiveSheet()->mergeCells('B13:B14');
                                $objPHPExcel->getActiveSheet()->getStyle('B13:B14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'name'      => 'Chaparral Pro',
                                            'size'      => '14'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objRichText = new PHPExcel_RichText();
                                $objPayable = $objRichText->createTextRun('Docente');
                                $objPayable->getFont()->setBold(true);
                                $objPHPExcel->getActiveSheet()->getCell('C13')->setValue($objRichText);
                                $objPHPExcel->getActiveSheet()->mergeCells('C13:C14');
                                $objPHPExcel->getActiveSheet()->getStyle('C13:C14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'name'      => 'Chaparral Pro',
                                            'size'      => '14'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objRichText = new PHPExcel_RichText();
                                $objPayable = $objRichText->createTextRun('Ciclo');
                                $objPayable->getFont()->setBold(true);
                                $objPHPExcel->getActiveSheet()->getCell('D13')->setValue($objRichText);
                                $objPHPExcel->getActiveSheet()->mergeCells('D13:D14');
                                $objPHPExcel->getActiveSheet()->getStyle('D13:D14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'name'      => 'Chaparral Pro',
                                            'size'      => '14'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objRichText = new PHPExcel_RichText();
                                $objPayable = $objRichText->createTextRun('Modalidad');
                                $objPayable->getFont()->setBold(true);
                                $objPHPExcel->getActiveSheet()->getCell('E13')->setValue($objRichText);
                                $objPHPExcel->getActiveSheet()->mergeCells('E13:E14');
                                $objPHPExcel->getActiveSheet()->getStyle('E13:E14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'name'      => 'Chaparral Pro',
                                            'size'      => '14'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );
                                
                                $objRichText = new PHPExcel_RichText();
                                $objPayable = $objRichText->createTextRun('Plan');
                                $objPayable->getFont()->setBold(true);
                                $objPHPExcel->getActiveSheet()->getCell('F13')->setValue($objRichText);
                                $objPHPExcel->getActiveSheet()->mergeCells('F13:F14');
                                $objPHPExcel->getActiveSheet()->getStyle('F13:F14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'name'      => 'Chaparral Pro',
                                            'size'      => '14'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );
                                $objRichText = new PHPExcel_RichText();
                                $objPayable = $objRichText->createTextRun('Programación y Organización');
                                $objPayable->getFont()->setBold(true);
                                $objPHPExcel->getActiveSheet()->getCell('G13')->setValue($objRichText);
                                $objPHPExcel->getActiveSheet()->mergeCells('G13:H13');
                                $objPHPExcel->getActiveSheet()->getStyle('G13:H13')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'name'      => 'Chaparral Pro',
                                            'size'      => '14'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );
                                $objPHPExcel->getActiveSheet()->setCellValue('G14','Calificación');

                                $objPHPExcel->getActiveSheet()->getStyle('G14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'size'      => '10'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objPHPExcel->getActiveSheet()->setCellValue('H14','Sub-Total');

                                $objPHPExcel->getActiveSheet()->getStyle('H14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'size'      => '10'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );



                                $objRichText = new PHPExcel_RichText();
                                $objPayable = $objRichText->createTextRun('Habilidades Motivación e');
                                $objPayable->getFont()->setBold(true);
                                $objPHPExcel->getActiveSheet()->getCell('I13')->setValue($objRichText);
                                $objPHPExcel->getActiveSheet()->mergeCells('I13:J13');
                                $objPHPExcel->getActiveSheet()->getStyle('I13:J13')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'name'      => 'Chaparral Pro',
                                            'size'      => '14'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );
                                $objPHPExcel->getActiveSheet()->setCellValue('I14','Calificación');
                                $objPHPExcel->getActiveSheet()->setCellValue('J14','Sub-Total');

                                $objPHPExcel->getActiveSheet()->getStyle('I14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'size'      => '10'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objPHPExcel->getActiveSheet()->getStyle('J14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'size'      => '10'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );


                                $objRichText = new PHPExcel_RichText();
                                $objPayable = $objRichText->createTextRun('Capacidad Académica');
                                $objPayable->getFont()->setBold(true);
                                $objPHPExcel->getActiveSheet()->getCell('K13')->setValue($objRichText);
                                $objPHPExcel->getActiveSheet()->mergeCells('K13:L13');
                                $objPHPExcel->getActiveSheet()->getStyle('K13:L13')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'name'      => 'Chaparral Pro',
                                            'size'      => '14'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );
                                
                                $objPHPExcel->getActiveSheet()->setCellValue('K14','Calificación');
                                $objPHPExcel->getActiveSheet()->setCellValue('L14','Sub-Total');

                                $objPHPExcel->getActiveSheet()->getStyle('K14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'size'      => '10'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objPHPExcel->getActiveSheet()->getStyle('L14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'size'      => '10'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objRichText = new PHPExcel_RichText();
                                $objPayable = $objRichText->createTextRun('Evaluación');
                                $objPayable->getFont()->setBold(true);
                                $objPHPExcel->getActiveSheet()->getCell('M13')->setValue($objRichText);
                                $objPHPExcel->getActiveSheet()->mergeCells('M13:N13');
                                $objPHPExcel->getActiveSheet()->getStyle('M13:N13')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'name'      => 'Chaparral Pro',
                                            'size'      => '14'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objPHPExcel->getActiveSheet()->setCellValue('M14','Calificación');
                                $objPHPExcel->getActiveSheet()->setCellValue('N14','Sub-Total');

                                    $objPHPExcel->getActiveSheet()->getStyle('M14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'size'      => '10'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objPHPExcel->getActiveSheet()->getStyle('N14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'size'      => '10'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );
                                 $objRichText = new PHPExcel_RichText();
                                $objPayable = $objRichText->createTextRun('Total');
                                $objPayable->getFont()->setBold(true);
                                $objPHPExcel->getActiveSheet()->getCell('O13')->setValue($objRichText);
                                $objPHPExcel->getActiveSheet()->mergeCells('O13:P13');
                                $objPHPExcel->getActiveSheet()->getStyle('O13:P13')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'name'      => 'Chaparral Pro',
                                            'size'      => '14'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objPHPExcel->getActiveSheet()->setCellValue('O14','Calificación');
                                $objPHPExcel->getActiveSheet()->setCellValue('P14','Total');

                                $objPHPExcel->getActiveSheet()->getStyle('O14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'size'      => '10'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                                $objPHPExcel->getActiveSheet()->getStyle('P14')->applyFromArray(
                                       array(
                                        'font'    => array(

                                            'bold'      => true,
                                            'size'      => '10'
                                        ),
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top'     => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                                );

                            $ind=0;
                            $cont=14;

                            for($i=0;$i<$na;$i++)
                                {
                                 /*
                                  * SELECT COUNT(DISTINCT(CODALU))
                                    FROM respuestas 
                                    WHERE CODCURR ='0096' AND CODCUR='03086' AND CODDOCE='CW0001 and CODMODALIDAD='01'
                                  */
                                 $re=$respuestas->listarSimple('COUNT(DISTINCT(CODALU))', 'CODCURR*CODCUR*CODDOCE*CODMODALIDAD', '=*=*=*=', $c[$i][4].'*'. $c[$i][5].'*'.$c[$i][6].'*'.$c[$i][7], 'AND*AND*AND*AND','','');
                                 $cont++;
                                 $ind++;
                                 $datos="$esc*$facu*$annio*$periodo*$ciclo*".$c[$i][4]."*".$c[$i][5].'*'.$c[$i][6].'*'.$c[$i][7];
                                 
                                 $promedios=traerPromedios($datos);
                                 $promediosD = explode("*", $promedios);
                                   
                                $objPHPExcel->getActiveSheet()->setCellValue('A'.$cont, $re[0][0]);
                                $objPHPExcel->getActiveSheet()->setCellValue('B'.$cont, utf8_encode($c[$i][0]));
                                $objPHPExcel->getActiveSheet()->setCellValue('C'.$cont, utf8_encode($c[$i][1])." ".utf8_encode($c[$i][2]));
                                $objPHPExcel->getActiveSheet()->setCellValue('D'.$cont, utf8_encode($c[$i][3]));
                                $objPHPExcel->getActiveSheet()->setCellValue('E'.$cont, utf8_encode($c[$i][8]));
                                $objPHPExcel->getActiveSheet()->setCellValue('F'.$cont, utf8_encode($c[$i][9]));

                                 $calificacion=calificarDocente($promediosD[0].'*1');
                                $objPHPExcel->getActiveSheet()->setCellValue('G'.$cont, $calificacion);
                                $objPHPExcel->getActiveSheet()->setCellValue('H'.$cont, $promediosD[0]);

                                $calificacion=calificarDocente($promediosD[1].'*2');
                                $objPHPExcel->getActiveSheet()->setCellValue('I'.$cont, $calificacion);
                                $objPHPExcel->getActiveSheet()->setCellValue('J'.$cont, $promediosD[1]);
                                 
                                 $calificacion=calificarDocente($promediosD[2].'*3');
                                 $objPHPExcel->getActiveSheet()->setCellValue('K'.$cont, $calificacion);
                                $objPHPExcel->getActiveSheet()->setCellValue('L'.$cont, $promediosD[2]);
                                 
                                 $calificacion=calificarDocente($promediosD[3].'*4');
                                $objPHPExcel->getActiveSheet()->setCellValue('M'.$cont, $calificacion);
                                $objPHPExcel->getActiveSheet()->setCellValue('N'.$cont, $promediosD[3]);
                                 
                                 $calificacion=calificarDocente($promediosD[4].'*5');
                                 $objPHPExcel->getActiveSheet()->setCellValue('O'.$cont, $calificacion);
                                $objPHPExcel->getActiveSheet()->setCellValue('P'.$cont, $promediosD[4]);


                                             $objPHPExcel->getActiveSheet()->getStyle('A'.$cont)->applyFromArray(
                                                array(
                                                    
                                                    'alignment' => array(
                                                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
                                                    ),
                                                    'borders' => array(
                                                        
                                                        'right'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'left'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'bottom'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        )
                                                    )
                                                )
                                        );
                                            $objPHPExcel->getActiveSheet()->getStyle('B'.$cont)->applyFromArray(
                                                array(
                                                    
                                                    'alignment' => array(
                                                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
                                                    ),
                                                    'borders' => array(
                                                       
                                                        'right'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'left'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'bottom'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        )
                                                    )
                                                )
                                        );
                                            $objPHPExcel->getActiveSheet()->getStyle('C'.$cont)->applyFromArray(
                                                array(
                                                    
                                                    'alignment' => array(
                                                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
                                                    ),
                                                    'borders' => array(
                                                        
                                                        'right'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'left'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'bottom'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        )
                                                    )
                                                )
                                        );
                                            $objPHPExcel->getActiveSheet()->getStyle('D'.$cont)->applyFromArray(
                                                array(
                                                    
                                                    'alignment' => array(
                                                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                                    ),
                                                    'borders' => array(
                                                        
                                                        'right'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'left'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'bottom'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        )
                                                    )
                                                )
                                        );
                                            $objPHPExcel->getActiveSheet()->getStyle('E'.$cont)->applyFromArray(
                                                array(
                                                    
                                                    'alignment' => array(
                                                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                                    ),
                                                    'borders' => array(
                                                        
                                                        'right'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'left'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'bottom'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        )
                                                    )
                                                )
                                        );
                                            $objPHPExcel->getActiveSheet()->getStyle('F'.$cont)->applyFromArray(
                                                array(
                                                    
                                                    'alignment' => array(
                                                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                                    ),
                                                    'borders' => array(
                                                        
                                                        'right'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'left'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'bottom'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        )
                                                    )
                                                )
                                        );
                                        $objPHPExcel->getActiveSheet()->getStyle('G'.$cont)->applyFromArray(
                                                array(
                                                    
                                                    'alignment' => array(
                                                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                                    ),
                                                    'borders' => array(
                                                       
                                                        'right'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'left'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        ),
                                                        'bottom'     => array(
                                                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                                        )
                                                    )
                                                )
                                        );

                                         $objPHPExcel->getActiveSheet()->getStyle('H'.$cont)->applyFromArray(
                                    array(
                                        
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                            );
                                $objPHPExcel->getActiveSheet()->getStyle('I'.$cont)->applyFromArray(
                                    array(
                                        
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                           
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                            );
                                $objPHPExcel->getActiveSheet()->getStyle('J'.$cont)->applyFromArray(
                                    array(
                                        
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                            );
                                $objPHPExcel->getActiveSheet()->getStyle('K'.$cont)->applyFromArray(
                                    array(
                                        
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                            );
                                $objPHPExcel->getActiveSheet()->getStyle('L'.$cont)->applyFromArray(
                                    array(
                                        
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                            );
                                $objPHPExcel->getActiveSheet()->getStyle('M'.$cont)->applyFromArray(
                                    array(
                                        
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                            );
                            $objPHPExcel->getActiveSheet()->getStyle('N'.$cont)->applyFromArray(
                                    array(
                                        
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                           
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                            );
                            $objPHPExcel->getActiveSheet()->getStyle('O'.$cont)->applyFromArray(
                                    array(
                                        
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                           
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                            );

                            $objPHPExcel->getActiveSheet()->getStyle('P'.$cont)->applyFromArray(
                                    array(
                                        
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                           
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                            );
                                } 

        }                      

                        /*    $objPHPExcel->getActiveSheet()->getStyle('G'.$cont)->applyFromArray(
                                    array(
                                        
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                           
                                            'right'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'left'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            ),
                                            'bottom'     => array(
                                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                            )
                                        )
                                    )
                            );
                                $objPHPExcel->getActiveSheet()->getStyle('G'.$cont)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                                $objPHPExcel->getActiveSheet()->getStyle('G'.$cont)->getFill()->getStartColor()->setARGB('DAEEFD');

                                if($re[0][0] == 0){
                                $objPHPExcel->getActiveSheet()->getStyle('G'.$cont)->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
                                }
                            }  
                               
                  }*/             


// Set column widths
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(4);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(6);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);

$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setAutoSize(true);



// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('Simple');

// Set page orientation and size
$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
// Set document security
$objPHPExcel->getSecurity()->setLockWindows(true);
$objPHPExcel->getSecurity()->setLockStructure(true);
$objPHPExcel->getSecurity()->setWorkbookPassword("shadow");


// Set sheet security
$objPHPExcel->getActiveSheet()->getProtection()->setPassword('shadow');
$objPHPExcel->getActiveSheet()->getProtection()->setSheet(true); // This should be enabled in order to enable any of the following!
$objPHPExcel->getActiveSheet()->getProtection()->setSort(true);
$objPHPExcel->getActiveSheet()->getProtection()->setInsertRows(true);
$objPHPExcel->getActiveSheet()->getProtection()->setFormatCells(true);
// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);


// Redirect output to a client’s web browser (Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="Reportes-'.$es[0][0].'-Ciclo-'.$ciclo.'.xlsx"');
header('Cache-Control: max-age=0');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;

function calificarDocente($promedio){
                 $datos = explode("*", $promedio);
                switch($datos[1])
                {
                    case 3:
                        if($datos[0]>=8 && $datos[0]<14.4){
                            $calificacion="Deficiente";
                            return $calificacion;
                        }
                        if($datos[0]>=14.4 && $datos[0]<20.8){
                            $calificacion="Regular";
                            return $calificacion;
                        }
                        if($datos[0]>=20.8 && $datos[0]<27.2){
                            $calificacion="Suficiente";
                            return $calificacion;
                        }
                        if($datos[0]>=27.2 && $datos[0]<33.6){
                            $calificacion="Bueno";
                            return $calificacion;
                        }
                        if($datos[0]>=33.6 && $datos[0]<=40){
                            $calificacion="Excelente";
                            return $calificacion;
                        }
                    break;
                
                
                    case 1: case 2: case 4:
                        
                         if($datos[0]>=4 && $datos[0]<7.2){
                            $calificacion="Deficiente";
                            return $calificacion;
                        }
                        if($datos[0]>=7.2 && $datos[0]<10.4){
                            $calificacion="Regular";
                            return $calificacion;
                        }
                        if($datos[0]>=10.4 && $datos[0]<13.6){
                            $calificacion="Suficiente";
                            return $calificacion;
                        }
                        if($datos[0]>=13.6 && $datos[0]<16.8){
                            $calificacion="Bueno";
                            return $calificacion;
                        }
                        if($datos[0]>=16.8 && $datos[0]<=20){
                            $calificacion="Excelente";
                            return $calificacion;
                        }
                        
                    break;
                
                    case 5:
                        if($datos[0]>=20 && $datos[0]<36){
                            $calificacion="Deficiente";
                            return $calificacion;
                        }
                        if($datos[0]>=36 && $datos[0]<52){
                            $calificacion="Regular";
                            return $calificacion;
                        }
                        if($datos[0]>=52 && $datos[0]<68){
                            $calificacion="Suficiente";
                            return $calificacion;
                        }
                        if($datos[0]>=68 && $datos[0]<84){
                            $calificacion="Bueno";
                            return $calificacion;
                        }
                        if($datos[0]>=84 && $datos[0]<=100){
                            $calificacion="Excelente";
                            return $calificacion;
                        }
                    break;
                        
                        
                }
                
            }
function traerPromedios($data){
                require_once("models/respuesta.php");
                require_once("models/pregunta.php");
                require_once("models/asignacion.php");
                require_once("models/oprespuesta.php");
                //var datos="esc="+a[1]+"&facu="+a[2]+"&annio="+a[3]+"&periodo="+a[4]+"&ciclo="+a[5]+"&codcurr="+a[6]+"&codcur="+a[7]
                //+"&coddoce="+a[8]+"&codmodalidad="+a[9]+"&boton=ReporteDocente";
                $datos = explode("*", $data);
               // print_r($datos);
                $esc=$datos[0];
                $facu=$datos[1];
                $annio=$datos[2];
                $periodo=$datos[3];
                $ciclo=$datos[4];
                $codcurr=$datos[5];
                $codcur=$datos[6];
                $coddoce=$datos[7];
                $codmodalidad=$datos[8];
                
                 $respuestas= new respuesta('', '', '','','','','','','');
                        $campo="re.`asignacione_id`*re.`pregunta_id`
                            *opres.`id`
                            *asig.`CODESC`*asig.CODFAC*asig.YYAKD*asig.CODPER*asig.YYCUR*
                            re.CODCURR*re.CODCUR*re.CODDOCE*re.`CODMODALIDAD`*pre.`dimensiones_id`";
                        
                        $operador="=*=*=*=*=*=*=*=*=*=*=*=*=";
                        $valor="asig.`id`*pre.`id`*re.`oprespuesta_id`*$esc*$facu*$annio*$periodo*$ciclo*$codcurr*$codcur*$coddoce*$codmodalidad*1";
                        $separador="AND*AND*AND*AND*AND*AND*AND*AND*AND*AND*AND*AND";
                        $resD1 = $respuestas->listar('opres.`valor`', $campo, $operador, $valor, $separador,'','');
                        $valor2="asig.`id`*pre.`id`*re.`oprespuesta_id`*$esc*$facu*$annio*$periodo*$ciclo*$codcurr*$codcur*$coddoce*$codmodalidad*2";
                        $resD2 = $respuestas->listar('opres.`valor`', $campo, $operador, $valor2, $separador,'','');
                        $valor3="asig.`id`*pre.`id`*re.`oprespuesta_id`*$esc*$facu*$annio*$periodo*$ciclo*$codcurr*$codcur*$coddoce*$codmodalidad*3";
                        $resD3 = $respuestas->listar('opres.`valor`', $campo, $operador, $valor3, $separador,'','');
                        $valor4="asig.`id`*pre.`id`*re.`oprespuesta_id`*$esc*$facu*$annio*$periodo*$ciclo*$codcurr*$codcur*$coddoce*$codmodalidad*4";
                        $resD4 = $respuestas->listar('opres.`valor`', $campo, $operador, $valor4, $separador,'','');
                        
                        $TresD1=count($resD1);
                        $al=$TresD1/4;
                        $i1=0;
                        $i2=4;
                        $i3=0;
                        $i4=8;
                        
                        $conta='1';
                                for($i=0;$i<$al;$i++){
                                    //TOTAL DIMENSION 1
                                     $dim=1;
                                     $total[$i][$dim]=0;
                                     for($j=$i1;$j<$i2;$j++){
                                         $total[$i][$dim]=$total[$i][$dim]+$resD1[$j][0];
                                    }

                                    //TOTAL DIMENSION 2
                                    $dim++;
                                    $total[$i][$dim]=0;
                                     for($j=$i1;$j<$i2;$j++){
                                         $total[$i][$dim]=$total[$i][$dim]+$resD2[$j][0];
                                    }
                                    
                                    //TOTAL DIMENSION 3
                                     $dim++;
                                    $total[$i][$dim]=0;
                                    for($k=$i3;$k<$i4;$k++){
                                         $total[$i][$dim]=$total[$i][$dim]+$resD3[$k][0];
                                    }
                                    
                                    //TOTAL DIMENSION 4
                                    $dim++;
                                    $total[$i][$dim]=0;
                                     for($j=$i1;$j<$i2;$j++){
                                         $total[$i][$dim]=$total[$i][$dim]+$resD4[$j][0];
                                    }
                                    
                                    //TOTAL POR ALUMNO
                                    $Ttotal[$i]=0;
                                    
                                    for($m=1;$m<=$dim;$m++){
                                        $Ttotal[$i]=$Ttotal[$i]+$total[$i][$m];
                                    }
                                    
                                    
                                    
                                     $i1=$i1+4;
                                     $i2=$i2+4;
                                     $i3=$i3+8;
                                     $i4=$i4+8;
                                     $conta++;
                                }
                                
                                 for($n=1;$n<=4;$n++){
                                    $suma=0;
                                    $prom[$n]=0;
                                    for($p=0;$p<$al;$p++){
                                        $suma=$total[$p][$n]+$suma;
                                        $prom[$n]=$suma/$al;
                                        $prom[$n]=round($prom[$n], 2);
                                    }
                                    
                                }
                                $sumaprom=0;
                                $Pprom=0;
                                for($p=0;$p<$al;$p++){
                                        $sumaprom=$Ttotal[$p]+$sumaprom;
                                        $Pprom=$sumaprom/$al;
                                        $Pprom=round($Pprom, 2);
                                    }
                                    
                               $promedios = $prom[1]."*".$prom[2]."*".$prom[3]."*".$prom[4]."*".$Pprom;
                             //  print_r($promedios)."</br>";
                                return $promedios;
                               
            }