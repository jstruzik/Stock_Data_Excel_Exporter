<?php
/** Mysql Connection */
require_once('mysqlManager.php');
/** PHPExcel */
include 'Classes/PHPExcel.php';
/** PHPExcel_Writer_Excel2007 */
include 'Classes/PHPExcel/Writer/Excel2007.php';

/**
 *	xlsxExport Class
 * An object that uses the PHPExcel reader and writer objects to parse and write new Excel files
 * @author Jacob Struzik
 *
 */
class xlsxExport{
   public $objReader;
   public $objPHPExcel;
   public $objWriter;
   public $db_conn_comp;
   public $row_counter;
    
   public function __construct(){
	   //our database holding company info
      $this->db_conn_comp = MySqlManager::CreateConnection('server', 'username', 'password', 'database');
      //our database holding table schema info
      $this->db_conn_info = MySqlManager::CreateConnection('server', 'username', 'password', 'information_schema');
      $this->generateReader();
      $this->row_counter = array();
   }

   /**
    * Inits our PHPExcel reader and loads a template Excel file
    */
   public function generateReader(){
      //create our reader
      $this->objReader = PHPExcel_IOFactory::createReader('Excel2007');
      $this->objReader->setLoadAllSheets();

      //load our template file into the reader object
      try {
         $this->objPHPExcel = $this->objReader->load("template.xlsx");
      }
      catch(Exception $e) {
         die('Error loading file: '.$e->getMessage());
      }
   }
   
   
   /**
    * Based on the selected companies and years, display the appropriate data
    * @param Array $tickers - desired company tickers
    * @param Int $years - desired range of years
    * @param String $startingFileDate - the initial date
    * @param boolean $export_all - flag to export all data
    */ 
   public function renderSelectedData($tickers,$years,$startingFileDate,$export_all){
      
      echo "Generating Excel Document, Please Wait...";

      if($export_all){
         $this->generateAll();
      }
      else{
         //our row counter sets our current row based on the excel template
         for($row_init = 0; $row_init < (2*$years)+1; $row_init++){
            $this->row_counter[$row_init] = 11;
            $this->row_counter[$row_init+$years] = 7;
            $this->row_counter[$row_init+$years+$years] = 6;
         }
          
         $stripped_date = explode("-",$startingFileDate);
         $current_year = $stripped_date[0];
         $request_trim = str_replace(" ", "",$tickers);
         $ticker_array = explode(",", $request_trim);
         $numTickers = count($ticker_array);
          
         //create new worksheets based on years
         $this->createWorksheetSelect($this->objPHPExcel,$years);
          
         $this->generateCompanyData($ticker_array,$years,$startingFileDate,$current_year,$numTickers);
         $this->generateExecutiveData($ticker_array,$years,$startingFileDate,$current_year,$numTickers);
         $this->generateBODData($ticker_array,$years,$startingFileDate,$current_year,$numTickers);
      }

      $this->writeExcelFile();
   }
    
   /**
    * Grab company data for specific tickers and write to Excel file
    * @param Array $ticker_array - desired company tickers
    * @param Int $years - desired range of years
    * @param String $startingFileDate - the initial date
    * @param Int $current_year - the current year we're iterating over
    * @param Int $numTickers - total number of tickers requested
    */
   public function generateCompanyData($ticker_array,$years,$startingFileDate,$current_year,$numTickers){

      //select our company comments table and fetch the columns in that table
      $table_name = "comp_comments";
      $cols_string = $this->fetch_comment_columns($table_name);

      for($i = 0; $i < $numTickers; $i++){
         $current_tick = $ticker_array[$i];
          
         //grab our selected companies
         $result = $this->db_conn_comp->FetchArray
         ("	SELECT *
               FROM `company_info`
               WHERE col0 = '$current_tick'
               AND col2 <= '$startingFileDate'
               ORDER BY `company_info`.`col2` ASC
               LIMIT 0 , $years");
          
         if(!empty($result)){
            foreach($result as $data_value){
               //grab our year
               $result_stripped_year = explode("-",$data_value["col2"]);
               $result_year = $result_stripped_year[0];
               //grab our current id for the row
               $current_id = $data_value["id_number"];
                
               //set our sheet index
               $this->objPHPExcel->setActiveSheetIndex($current_year-$result_year);
                
               //actually set the values in the excel file
               for($j=0;$j<183;$j++){
                  $cell = $this->numtochars($j+1) . $this->row_counter[$current_year-$result_year];
                  if($j!=0 and $j!=1 and $j!=2){
                     $split_data = explode('|||',$data_value["col$j"]);
                     $old_row = $split_data[0];
                     $new_data = $split_data[1];
                      
                     #Check if it's a formula
                     if(substr($new_data,0,1) == '='){
                        #Now check if it's strictly a numeral function or one with cell values
                        if($this->checkFormula($new_data)){
                        $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
                     }
                     #Replace all row #'s with current row.
                     else{
                        $new_data = $this->parseFormula($new_data,$this->row_counter[$current_year-$result_year],$old_row);
                     }
                     }
                     $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
                  }
                  else{
                     $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $data_value["col$j"]);
                  }
               }
                
               $this->fetch_comments($this->objPHPExcel,$this->row_counter[$current_year-$result_year],$current_tick,$startingFileDate,$years,$table_name,$current_year,$current_id, $cols_string);
            }
            $this->row_counter[$current_year-$result_year]++;
         }
      }
   }

   /**
    * Grab Executive data for specific tickers and write to Excel file
    * @param Array $ticker_array - desired company tickers
    * @param Int $years - desired range of years
    * @param String $startingFileDate - the initial date
    * @param Int $current_year - the current year we're iterating over
    * @param Int $numTickers - total number of tickers requested
    */
   public function generateExecutiveData($ticker_array,$years,$startingFileDate,$current_year,$numTickers){
       
      //select our company comments table and fetch the columns in that table
      $table_name = "exec_comments";
      $cols_string = $this->fetch_comment_columns($table_name);

      for($i = 0; $i < $numTickers; $i++){
         $current_tick = $ticker_array[$i];
          
         $whichExecs = $this->determineExecutives();
         $queryResult = $this->execQuery($whichExecs, $current_tick, $startingFileDate, $years);
          
         //grab our selected executives
         $result = $this->db_conn_comp->FetchArray($queryResult);

         if(!empty($result)){
            foreach($result as $data_value){
               //grab our year
               $result_stripped_year = explode("-",$data_value["col4"]);
               $result_year = $result_stripped_year[0];
               //grab our current id for the row
               $current_id = $data_value["id_number"];
                
               //set our sheet index
               $this->objPHPExcel->setActiveSheetIndex($current_year-$result_year+$years);
                
               //actually set the values in the excel file
               for($j=0;$j<105;$j++){
                  $cell = $this->numtochars($j+1) . $this->row_counter[$current_year-$result_year+(2*$years)];
                  if($j!=2 and $j!=3 and $j!=4 and $j!=10){
                     $split_data = explode('|||',$data_value["col$j"]);
                     $old_row = $split_data[0];
                     $new_data = $split_data[1];

                     #Check if it's a formula
                     if(substr($new_data,0,1) == '='){
                        #Now check if it's strictly a numeral function or one with cell values
                        if($this->checkFormula($new_data)){
                        $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
                     }
                     #Replace all row #'s with current row.
                     $new_data = $this->parseFormula($new_data,$this->row_counter[$current_year-$result_year+(2*$years)],$old_row);
                     }
                     $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
                  }
                  else{
                     $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $data_value["col$j"]);
                  }
               }
               $this->fetch_comments($this->objPHPExcel,$this->row_counter[$current_year-$result_year+(2*$years)],$current_tick,$startingFileDate,$years,$table_name,$current_year,$current_id, $cols_string);
            }
            $this->row_counter[$current_year-$result_year+(2*$years)]++;
         }
      }
   }

   /**
    * Grab board of directors data for specific tickers and write to Excel file
    * @param Array $ticker_array - desired company tickers
    * @param Int $years - desired range of years
    * @param String $startingFileDate - the initial date
    * @param Int $current_year - the current year we're iterating over
    * @param Int $numTickers - total number of tickers requested
    */
   public function generateBODData($ticker_array,$years,$startingFileDate,$current_year,$numTickers){
       
      //select our company comments table and fetch the columns in that table
      $table_name = "bod_comments";
      $cols_string = $this->fetch_comment_columns($table_name);
       
      for($i = 0; $i < $numTickers; $i++){
         $current_tick = $ticker_array[$i];

         //grab our selected companies
         $result = $this->db_conn_comp->FetchArray
         ("	SELECT *
               FROM `bod_info`
               WHERE col1 = '$current_tick'
               AND col2 <= '$startingFileDate'
               ORDER BY `bod_info`.`col2` ASC
               LIMIT 0 , $years");

         if(!empty($result)){
            foreach($result as $data_value){
               //grab our year
               $result_stripped_year = explode("-",$data_value["col2"]);
               $result_year = $result_stripped_year[0];
               //grab our current id for the row
               $current_id = $data_value["id_number"];
                

               //set our sheet index
               $this->objPHPExcel->setActiveSheetIndex($current_year-$result_year+(2*$years));

               //actually set the values in the excel file
               for($j=0;$j<100;$j++){
                  $cell = $this->numtochars($j+1) . $this->row_counter[$current_year-$result_year+(3*$years)];
                  if($j!=0 and $j!=1 and $j!=2){
                     $split_data = explode('|||',$data_value["col$j"]);
                     $old_row = $split_data[0];
                     $new_data = $split_data[1];

                     #Check if it's a formula
                     if(substr($new_data,0,1) == '='){
                        #Now check if it's strictly a numeral function or one with cell values
                        if($this->checkFormula($new_data)){
                        $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
                     }
                     #Replace all row #'s with current row.
                     else{
                        $new_data = $this->parseFormula($new_data,$this->row_counter[$current_year-$result_year+(3*$years)],$old_row);
                     }
                     }
                     $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
                  }
                  else{
                     $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $data_value["col$j"]);
                  }
               }
               $this->fetch_comments($this->objPHPExcel,$this->row_counter[$current_year-$result_year+(3*$years)],$current_tick,$startingFileDate,$years,$table_name,$current_year,$current_id, $cols_string);
            }
            $this->row_counter[$current_year-$result_year+(3*$years)]++;
         }
      }
   }

   /**
    * Grab all possible data and write to Excel file
    */
   public function generateAll(){
      $comp = $this->db_conn_comp->FetchArray("SELECT MIN(col2)
            FROM company_info");
      $exec = $this->db_conn_comp->FetchArray("SELECT MIN(col4)
            FROM exec_info");
      $bod = $this->db_conn_comp->FetchArray("SELECT MIN(col2)
            FROM bod_info");
      $earliest_comp = explode('-',$comp[0]["MIN(col2)"]);
      $earliest_exec = explode('-',$exec[0]["MIN(col4)"]);
      $earliest_bod = explode('-',$bod[0]["MIN(col2)"]);
      $current_year = date("Y");

      $max_years = max(array(($current_year-$earliest_comp[0]),($current_year-$earliest_exec[0]),($current_year-$earliest_bod[0])));

      $this->createWorksheetSelect($this->objPHPExcel,$max_years);

      //our row counter sets our current row based on the excel template
      for($row_init = 0; $row_init < (2*$max_years)+1; $row_init++){
         $this->row_counter[$row_init] = 11;
         $this->row_counter[$row_init+$max_years] = 7;
         $this->row_counter[$row_init+$max_years+$max_years] = 6;
      }

      /////COMPANY ALL\\\\\\\
      //select our company comments table and fetch the columns in that table
      $table_name = "comp_comments";
      $cols_string = $this->fetch_comment_columns($table_name);
      
      $result = $this->db_conn_comp->FetchArray
      ("SELECT *
            FROM `company_info`");

      if(!empty($result)){
         foreach($result as $data_value){
            //grab our year
            $result_stripped_year = explode("-",$data_value["col2"]);
            $result_year = $result_stripped_year[0];
            //grab our current id for the row
            $current_id = $data_value["id_number"];

            //set our sheet index
            $this->objPHPExcel->setActiveSheetIndex($current_year-$result_year);
             
            //actually set the values in the excel file
            for($j=0;$j<183;$j++){
               $cell = $this->numtochars($j+1) . $this->row_counter[$current_year-$result_year];
               if($j!=0 and $j!=1 and $j!=2){
                  $split_data = explode('|||',$data_value["col$j"]);
                  $old_row = $split_data[0];
                  $new_data = $split_data[1];
                   
                  #Check if it's a formula
                  if(substr($new_data,0,1) == '='){
                     #Now check if it's strictly a numeral function or one with cell values
                     if($this->checkFormula($new_data)){
                     $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
                  }
                  #Replace all row #'s with current row.
                  else{
                     $new_data = $this->parseFormula($new_data,$this->row_counter[$current_year-$result_year],$old_row);
                  }
                  }
                  $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
               }
               else{
                  $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $data_value["col$j"]);
               }
            }
            $this->fetch_all_comments($this->objPHPExcel,$this->row_counter[$current_year-$result_year],$table_name,$current_id, $cols_string);
            $this->row_counter[$current_year-$result_year]++;
         }
      }
      /////EXECUTIVE ALL\\\\\\\
      $table_name = "exec_comments";
      $cols_string = $this->fetch_comment_columns($table_name);
      
      $result = $this->db_conn_comp->FetchArray
      ("SELECT *
            FROM `exec_info`");

      if(!empty($result)){
         foreach($result as $data_value){
            //grab our year
            $result_stripped_year = explode("-",$data_value["col4"]);
            $result_year = $result_stripped_year[0];
            //grab our current id for the row
            $current_id = $data_value["id_number"];

            //set our sheet index
            $this->objPHPExcel->setActiveSheetIndex($current_year-$result_year+$max_years);
             
            //actually set the values in the excel file
            for($j=0;$j<105;$j++){
               $cell = $this->numtochars($j+1) . $this->row_counter[$current_year-$result_year+(2*$max_years)+1];
               if($j!=2 and $j!=3 and $j!=4 and $j!=10){
                  $split_data = explode('|||',$data_value["col$j"]);
                  $old_row = $split_data[0];
                  $new_data = $split_data[1];
                   
                  #Check if it's a formula
                  if(substr($new_data,0,1) == '='){
                     #Now check if it's strictly a numeral function or one with cell values
                     if($this->checkFormula($new_data)){
                     $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
                  }
                  #Replace all row #'s with current row.
                  else{
                     $new_data = $this->parseFormula($new_data,$this->row_counter[$current_year-$result_year+(2*$max_years)+1],$old_row);
                  }
                  }
                  $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
               }
               else{
                  $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $data_value["col$j"]);
               }
            }
            $this->fetch_all_comments($this->objPHPExcel,$this->row_counter[$current_year-$result_year+(2*$max_years)+1],$table_name,$current_id, $cols_string);
            $this->row_counter[$current_year-$result_year+(2*$max_years)+1]++;
         }
      }
      /////BOD ALL\\\\\\\
      $table_name = "bod_comments";
      $cols_string = $this->fetch_comment_columns($table_name);
      
      $result = $this->db_conn_comp->FetchArray
      ("SELECT *
            FROM `bod_info`");

      if(!empty($result)){
         foreach($result as $data_value){
            //grab our year
            $result_stripped_year = explode("-",$data_value["col2"]);
            $result_year = $result_stripped_year[0];
            //grab our current id for the row
            $current_id = $data_value["id_number"];

            //set our sheet index
            $this->objPHPExcel->setActiveSheetIndex($current_year-$result_year+($max_years*2));
             
            //actually set the values in the excel file
            for($j=0;$j<183;$j++){
               $cell = $this->numtochars($j+1) . $this->row_counter[$current_year-$result_year+(3*$max_years)+1];
               if($j!=0 and $j!=1 and $j!=2){
                  $split_data = explode('|||',$data_value["col$j"]);
                  $old_row = $split_data[0];
                  $new_data = $split_data[1];
                   
                  #Check if it's a formula
                  if(substr($new_data,0,1) == '='){
                     #Now check if it's strictly a numeral function or one with cell values
                     if($this->checkFormula($new_data)){
                     $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
                  }
                  #Replace all row #'s with current row.
                  else{
                     $new_data = $this->parseFormula($new_data,$this->row_counter[$current_year-$result_year+(3*$max_years)+1],$old_row);
                  }
                  }
                  $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $new_data);
               }
               else{
                  $this->objPHPExcel->getActiveSheet()->SetCellValue("$cell", $data_value["col$j"]);
               }
            }
            $this->fetch_all_comments($this->objPHPExcel,$this->row_counter[$current_year-$result_year+(3*$max_years)+1],$table_name,$current_id, $cols_string);
            $this->row_counter[$current_year-$result_year+(3*$max_years)+1]++;
         }
      }
   }

   /**
    * Builds Excel file from data
    */
   public function writeExcelFile(){
      // Save Excel 2007 file
      $this->objWriter = new PHPExcel_Writer_Excel2007($this->objPHPExcel);
      $this->objWriter->setPreCalculateFormulas(false);
      $this->objWriter->save(str_replace('.php', '.xlsx', __FILE__));
      echo"Complete!";
      
   }

   /**
    * Grab all of the data that contain Excel comments
    */
   private function fetch_comment_columns($table_name){
      //grab all of our column names
      $num_col= $this->db_conn_info->FetchArray("SELECT column_name FROM columns
            WHERE table_name = '$table_name'");
      //initiate our column name string
      $cols_string = '';

      foreach($num_col as $column){
         $omitted_columns = array('id_number','company','date_info','executive');
         if(!in_array($column['column_name'],$omitted_columns)){
            $cols_string .= trim($column['column_name']).",";
         }
      }
       
      //remove the last comma in the string
      $cols_string = substr($cols_string,0,-1);
      return $cols_string;
   }
    
   private function fetch_comments($objPHPExcel,$current_row,$tick,$startingFileDate,$years,$table_name,$current_year,$current_id,$cols_string){
       
      //search comments using desired date and company name and id number
      $result = $this->db_conn_comp->FetchArray("SELECT $cols_string FROM $table_name WHERE company='$tick'
            AND date_info<='$startingFileDate' AND id_number='$current_id'
            ORDER BY $table_name.`date_info`
            ASC LIMIT 0, $years");

      foreach($result as $comm_search){
         foreach($comm_search as $search_value){
            if($search_value!='' && $search_value != 'NULL' && $search_value != null){
               //split our values by the |
               $split_value = explode("|", $search_value);
               for($i=0;$i<count($split_value);$i+=2){
                  //Since we need to align the comments to the rows of the selected data, strip the numbers and add on the current row (we won't do this in export all)
                  $comm_cell = $this->remove_numbers($split_value[$i]).($current_row);
                  $comm_value = $split_value[$i+1];
                  //write our comments to the proper location
                  $this->objPHPExcel->getActiveSheet()->getComment($comm_cell)->getText()->createTextRun($comm_value);
               }
            }
         }
      }
   }
   
   private function fetch_all_comments($objPHPExcel,$current_row,$table_name,$current_id,$cols_string){
   
      $sql = "SELECT $cols_string FROM $table_name WHERE id_number='$current_id' ORDER BY $table_name.`date_info` ASC";
      //search comments using desired date and company name and id number
      $result = $this->db_conn_comp->FetchArray("SELECT $cols_string FROM $table_name WHERE id_number='$current_id'
            ORDER BY $table_name.`date_info`
            ASC");

      foreach($result as $comm_search){
         foreach($comm_search as $search_value){
            if($search_value!='' && $search_value != 'NULL' && $search_value != null){
               //split our values by the |
               $split_value = explode("|", $search_value);
               for($i=0;$i<count($split_value);$i+=2){
                  //Since we need to align the comments to the rows of the selected data, strip the numbers and add on the current row (we won't do this in export all)
                  $comm_cell = $this->remove_numbers($split_value[$i]).($current_row);
                  $comm_value = $split_value[$i+1];
                  //write our comments to the proper location
                  $this->objPHPExcel->getActiveSheet()->getComment($comm_cell)->getText()->createTextRun($comm_value);
               }
            }
         }
      }
   }

   /**
    * Generate Excel worksheets based on the number of years
    * @param objPHPExcel $objPHPExcel - our PHPExcel object
    * @param Int $years - the year range
    */
   private function createWorksheetSelect ($objPHPExcel,$years){
      $comp_sheet = $objPHPExcel->setActiveSheetIndex(0)->copy();
      $exec_sheet = $objPHPExcel->setActiveSheetIndex(1)->copy();
      $bod_sheet = $objPHPExcel->setActiveSheetIndex(2)->copy();
       
      //removes the exec and BOD sheets for the moment
      $objPHPExcel->removeSheetByIndex(2);
      $objPHPExcel->removeSheetByIndex(1);
      $objPHPExcel->getSheet(0)->freezePane('C11');
       
      $counter = 0;
      $counter2 = 0;
       
      for($num=1;$num<$years;$num++){
         $new = clone $comp_sheet;
         $objPHPExcel->addSheet($new,$num);
         $objPHPExcel->getSheet($num)->setTitle("Company DB(-" . ($num) . ")");
         $objPHPExcel->getSheet($num)->freezePane('C11');
         $counter++;
      }
       
      for($num=1;$num<=$years;$num++){
         $new2 = clone $exec_sheet;
         $objPHPExcel->addSheet($new2,($num+$counter));
         if($num == 1){
            $objPHPExcel->getSheet(($num+$counter))->setTitle("ExecDB");
            $objPHPExcel->getSheet($num+$counter)->freezePane('L7');
         }else{
            $objPHPExcel->getSheet(($num+$counter))->setTitle("ExecDB(-" . ($num-1) . ")");
            $objPHPExcel->getSheet($num+$counter)->freezePane('L7');
         }
         $counter2++;
      }
       
      for($num=1;$num<=$years;$num++){
         $new3 = clone $bod_sheet;
         $objPHPExcel->addSheet($new3,($num+$counter+$counter2));
         if($num == 1){
            $objPHPExcel->getSheet(($num+$counter+$counter2))->setTitle("BOD");
            $objPHPExcel->getSheet($num+$counter+$counter2)->freezePane('I7');
         }else{
            $objPHPExcel->getSheet(($num+$counter+$counter2))->setTitle("BOD(-" . ($num-1) . ")");
            $objPHPExcel->getSheet($num+$counter+$counter2)->freezePane('I7');
         }
      }
   }
   
   /**
    * A function to check if the passed cell is a formula
    * @param String $formula - the given formula
    * @return - true if formula is valid
    */
   private function checkFormula($formula){
      $possible_operators = array('+','-','=','*','/',',','(',')','[',']','{','}','#','<','>');
      $numbers = array('0','1','2','3','4','5','6','7','8','9');
      
      for($i=0;$i<strlen($formula);$i++){
         if(!in_array($formula[$i],$possible_operators) && !in_array($formula[$i],$numbers)){
            return false;
         }
      }
      
      return true;
   }
   
   /**
    * A function to parse the formula to avoid PHPExcels attempt to evaluate functions
    * @param String $formula - the formula to be parsed
    * @param Int $current_row - the current row in the iteration
    * @param Int $old_row - the row the formula used to exist on so that an offset may be generated
    * when new worksheets are created
    */
   private function parseFormula($formula,$current_row,$old_row){
      $possible_operators = array('+','-','=','*','/',',','(',')','[',']','{','}','#','<','>');
      $numbers = array('0','1','2','3','4','5','6','7','8','9');
      $full_token = array();
      $row_tokens = array();
      $new_row = 0;
      $offset = $current_row - $old_row;
      $formula_count = strlen($formula);
      
      for($i=0;$i<=$formula_count;$i++){
         if($i == 0){
            array_push($full_token,'=');
         }
         if(in_array($formula[$i],$possible_operators)){
            for($j=$i+1;$j<=$formula_count;$j++){
               if(!in_array($formula[$j],$numbers)){
                  if(in_array($formula[$j],$possible_operators)){
                     array_push($full_token,$formula[$j]);
                     continue 2;
                  }
                  else{
                     array_push($full_token,$formula[$j]);
                     continue;
                  }
               }
               #Make sure we're grabbing a valid row
               else{
                  #This pushes a number and makes sure to grab ints with more than one digit
                  $number = array();
                  array_push($number,$formula[$j]);
                  $start = $j-1;
                  for($k=$j+1;$k<=$formula_count;$k++){
                     if(in_array($formula[$k],$numbers)){
                        array_push($number,$formula[$k]);
                     }
                     else{
                        array_push($row_tokens,implode('',$number));
                        $j = $k-1;
                        break;
                     }
                  }
                  $new_row = implode('',$row_tokens);
                  if(preg_match('/[A-Z]/',$formula[$start])){
                     $row = $new_row+$offset;
                  }
                  else{
                     $row = $new_row;
                  }
                  array_push($full_token,$row);
                  $row_tokens = array();
               }
            }
         }
      }
      return join($full_token);
   }
   
   /**
    * Defaults values when determining excutive positions
    */
   private function determineExecutives(){
      $execs = array('Top_5'=>true,'CEO'=>false, 'CFO'=>false, 'COO'=>false);
      
      if(isset($_POST['CEO']) && $_POST['CEO'] == 'CEO'){
         $execs['CEO'] = true;
         $execs['Top_5'] = false;
      }
      if(isset($_POST['CFO']) && $_POST['CFO'] == 'CFO'){
         $execs['CFO'] = true;
         $execs['Top_5'] = false;
      }
      if(isset($_POST['COO']) && $_POST['COO'] == 'COO'){
         $execs['COO'] = true;
         $execs['Top_5'] = false;
      }
      return $execs;
   }
   
   /**
    * A customized query to determine which executive positions to grab
    */
   private function execQuery($executives, $ticker, $startingFileDate, $years){
      $CEO = '';
      $CFO = '';
      $COO = '';
   
      $endFileDate = $startingFileDate - $years+1;
   
      if($executives['Top_5'] == true){
         echo "get results for top 5 <br>";
         $result = ("	SELECT *
                                 FROM `exec_info`
                                 WHERE col2 = '$ticker'
                                 AND col4 BETWEEN '$endFileDate'
                                 AND	'$startingFileDate'
                                 ORDER BY `exec_info`.`col4` ASC
                                 ");
      }
      else{
         $firstEntry = true;
         if($executives['CEO'] == true){
            $CEO = "col13 = 'CEO'";
            $firstEntry = false;
         }
         if($executives['CFO'] == true){
            if($firstEntry == true){
               $CFO = "col13 = 'CFO'";
               $firstEntry = false;
            }
            else{
               $CFO = "OR col13 = 'CFO'";
            }
         }
         if($executives['COO'] == true){
            if($firstEntry == true){
               $COO = "col13 = 'COO'";
               $firstEntry = false;
            }
            else{
               $COO = "OR col13 = 'COO'";
            }
         }

         $result = ("	SELECT *
                                 FROM `exec_info`
							            WHERE col2 = '$ticker'
         								AND col4 BETWEEN '$endFileDate'
         											AND	'$startingFileDate'
         								AND (	$CEO
         										$COO
         										$CFO	)
         								ORDER BY `exec_info`.`col4` DESC
								         ");
   	}
   	return $result;
   }
   
   /**
    * Simple function to convert row numbers to excel characters
    */
   private function numtochars($num,$start=65,$end=90)
   {
   
      $num = abs($num);
      $str = "";
      $cache = 26;
      while($num != 0)
      {
         $temp = $num%$cache;
         if($temp == 0){
            $temp = 26;
         }
         $str = chr($temp+$start-1).$str;
         $num = ($num-$temp)/$cache;
      }
      return $str;
   }
   
   /**
    * Simple function to strip out all numbers from a string
    */
   private function remove_numbers($string) {
      $vowels = array("1", "2", "3", "4", "5", "6", "7", "8", "9", "0", " ");
      $string = str_replace($vowels, '', $string);
      return $string;
   }
}
?>
