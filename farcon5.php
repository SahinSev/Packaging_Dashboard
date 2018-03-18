<?php
//Include DB- connect
include '../Jensovic/dbconnect.php';

//  Include PHPExcel_IOFactory
include 'PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';


// Input filename decleration
$inputFileName = './tbl_rohdaten_verpackung.xlsx';

       $objPHPExcel = PHPExcel_IOFactory::load($inputFileName);
       $objWorksheet = $objPHPExcel->getActiveSheet();

       $zeile = 0;

       $insert_code_1 = " insert into `auswertung_tbl_verpackung` "; # 3 year review
       
       $insert_code_1.= " ( ";

       $insert_code_1.= "    `Auftrag_ID`           , ";
       $insert_code_1.= "    `Anlage`               , ";
       $insert_code_1.= "    `FAUF`                 , ";
       $insert_code_1.= "    `Istmenge`             , ";
       $insert_code_1.= "    `Geni_NR`              , ";
       $insert_code_1.= "    `Vertriebstext`        , ";
       $insert_code_1.= "    `Produktgruppe`        , ";
       $insert_code_1.= "    `AHG`                  , ";
       $insert_code_1.= "    `Land-Bez`             , ";
       $insert_code_1.= "    `Konzentration`        , ";
       $insert_code_1.= "    `Einheit`              , ";
       $insert_code_1.= "    `IstZeitSAP_Fix`       , ";
       $insert_code_1.= "    `IstZeitSap_Variabel`  , ";
       $insert_code_1.= "    `AuftragStart_MES`     , ";
       $insert_code_1.= "    `AuftragEnde_MES`        ";


       $insert_code_1.= " ) ";
       
       $insert_code_1.= "  values ";
       
       $insert_code_1.= " ( ";

       foreach ($objWorksheet->getRowIterator() as $row) {

           $zeile++;

           $Fehler_Daten = 0;

           $cellIterator = $row->getCellIterator();
           $cellIterator->setIterateOnlyExistingCells(false);

           $spalte = 0;

           foreach ($cellIterator as $cell) {

               $spalte++;

               $wert = $cell->getValue();
               #$wert = trim(iconv("UTF-8", "ISO-8859-1", $wert));
               
               $wert = trim($wert);

               #echo"<br>Zeile $zeile, Spalte $spalte: $wert<br>";

               $Arr_Values[$zeile][$spalte] = $wert;

               $insert_code_2.= $Arr_Values[$zeile][$spalte] . ", ";             

           }
           
          $insert_code_2.= " "; # 3 year review
          
         }

          $insert_code_2 = rtrim($insert_code_2,", ");
          
          $insert_code_2.= ")";

          // Auslesen des Arrays ' & , ergänzen und linearisieren anschließend in $arra_auslesen

          $insert_code_4 = "";

          $arr_auslesen = array();

          $i = 0;

          foreach ($Arr_Values As $array) {
            
            $Counter = 0;
            
            $i++; 
            
            $spalte = 0;
            
            while ($Counter < 13) {               
                foreach ($array AS $arr) {
                    $arr = "'" . $arr . "'";
                    #echo '<pre>'; print ($arr) . " " . $Counter ; echo '</pre>';

                    $insert_code_4 = $insert_code_4 . $arr . ", ";

                    $Counter++;
                }
                
              }
              $insert_code_4 = rtrim($insert_code_4,", ") . ") ";
              #echo '<pre>'; print ($insert_code_4) ; echo '</pre>';
              $arr_auslesen[$i] = $insert_code_4 ;

              #echo '<pre>'; print $arr_auslesen[$i] . $i ; echo '</pre>';

              $insert_code_4 = "";
            }

          unset($arr_auslesen[1]);
          $arr_auslesen = array_values($arr_auslesen);
          
          #echo '<pre>'; print $arr_auslesen[1]  ; echo '</pre>';

            
            foreach ($arr_auslesen As $arr) {

              $sql = $insert_code_1 . $arr; 

              echo '<pre>'; print ($sql) ; echo '</pre>';
                if ($mysql_conn->query($sql) === TRUE) 
                {
                  #$message = "Geodaten wurde erfolgreich eingetragen!";
                  #echo "<script type='text/javascript'>alert('$message');</script>";
                  #echo "Account_Name wurde erfolgreich geändert!";
                } else 
                {
                  echo "Error: " . $sql . "<br>" . $mysql_conn->error;
                }
                
            
            }

            $mysql_conn->close();
         
?>
