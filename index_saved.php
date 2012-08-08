<?php
/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2012 PHPExcel
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
 * @copyright  Copyright (c) 2006 - 2012 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.7.7, 2012-05-19
 */


set_time_limit(0);

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', 'On');
//ini_set('memory_limit', '-1');

date_default_timezone_set('Europe/Budapest');

require_once 'simple_html_dom.php';
require_once 'phpexcel/Classes/PHPExcel/IOFactory.php';
require_once 'phpexcel/Classes/PHPExcel.php';


$inputFileName = 'sampleData/teszt.xls';
$objPHPExcel = PHPExcel_IOFactory::load($inputFileName);
$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Set document properties
$objPHPExcel->getProperties()->setCreator("Brandlift")
                             ->setLastModifiedBy("Brandlift")
                             ->setTitle("Brandlift Document")
                             ->setSubject("Brandlift Document");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', 'Eredeti link')
            ->setCellValue('B1', 'Title')
            ->setCellValue('C1', 'Created')
            ->setCellValue('D1', 'Type')
            ->setCellValue('E1', 'Members')
            ->setCellValue('F1', 'Owner')
            ->setCellValue('G1', 'Website')
            ->setCellValue('H1', 'Content');

$counter = 1;
foreach($sheetData as $data) {

    if ($counter < 5) {

        $link = $data['A'];

        if ($link != "") {
            
            $counter++;
            
            $is_there_S = false;
            
            // van-e benne .S.
            if (strpos($link, '.S.') > 0 ) {
                
                $link = substr($link, 0, strpos($link, '.S.'));
                $is_there_S = true;
            }
            
            // ha van benne &gid=
            if (strpos($link, '&gid=') > 0 ) {
                
                
                $gid = substr($link, strpos($link, '&gid='));
                $gid = substr($gid, 1,11);
                $link = 'http://www.linkedin.com/groups?'.$gid;
            }
            
            $html = file_get_html($link);
            
            if ($is_there_S) foreach($html->find('.group-name') as $element) $title = strip_tags ($element->outertext);
            else foreach($html->find('.group-name') as $element) $title = strip_tags ($element->outertext);
            
            foreach($html->find('.anet-navbox ul li') as $li_element) {
                
                $li = trim( strip_tags($li_element));
                $exploded_li = explode(":", $li);
                
                $key = $exploded_li[0];
                $value = "";
                foreach ($exploded_li as $curr_key => $curr_value) {
                    
                    if ($curr_key > 0) $value .= $curr_value;
                }
                
                if ($key == 'Created') $created = $value;
                if ($key == 'Type') $type = $value;
                if ($key == 'Members') $members = $value;
                if ($key == 'Owner') $owner = $value;
                if ($key == 'Website') $website = $value;
            }
            
            foreach($html->find('#content .groups-upsell-SEO') as $element) $content = strip_tags ($element->outertext);
             
            // Add some data
            $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('A'.$counter, $link)
                        ->setCellValue('B'.$counter, $title)
                        ->setCellValue('C'.$counter, $created)
                        ->setCellValue('D'.$counter, $type)
                        ->setCellValue('E'.$counter, $members)
                        ->setCellValue('F'.$counter, $owner)
                        ->setCellValue('G'.$counter, $website)
                        ->setCellValue('H'.$counter, $content);

            // Set active sheet index to the first sheet, so Excel opens this as the first sheet
            $objPHPExcel->setActiveSheetIndex(0);
        }
    }
}

// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="01simple.xls"');
header('Cache-Control: max-age=0');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
            
            