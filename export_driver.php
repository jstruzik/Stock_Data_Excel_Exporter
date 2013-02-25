<?php
require_once('xlsxExport.php');

$export_class = new xlsxExport();

if(isset($_POST['selection'])){
   echo'<div class="debug_info">';
   $export_class->renderSelectedData($_POST["ticker"],$_POST["years"],inToSqlDate($_POST["fileDate"]),false);
   echo'</div>';
}
else if(isset($_POST['all'])){
   echo'<div class="debug_info">';
   $export_class->renderSelectedData('','','',true);
   echo'</div>';
}


function inToSqlDate($inDate){
   //if no date specified, return todays date
   if($inDate == ""){
      $requestDate = date("Y-m-d");
      return $requestDate;
   }else{
      //convert date entered into mySQL style
      $date = explode("/",$inDate);
      $year = $date[2];
      $month = $date[0];
      $day = $date[1];
      //Convert to mySQL year syntax
      $yearLength = strlen($year);
      switch($yearLength){
         case 1:
            die("You did not enter a valid year type <br>");
         case 2:
            $year = "20" . $year;
            //echo $year . "<br>";
            break;
         case 3:
            die("You did not enter a valid year type <br>");
         case 4:
            break;
         default:
            die("You did not enter a valid year type <br>");
      }
      //Convert to mySQL month syntax
      $monthLength = strlen($month);
      switch($monthLength){
         case 1:
            $month = "0" . $month;
            //echo $month . "<br>";
            break;
         case 2:
            break;
         default:
            die("You did not enter a valid month type <br>");
      }
      //Convert to mySQL day syntax
      $dayLength = strlen($day);
      switch($dayLength){
         case 1:
            $day = "0" . $day;
            //echo $day . "<br>";
            break;
         case 2:
            break;
         default:
            die("You did not enter a valid day type <br>");
      }
      //Check to make sure that it is a valid date
      //intval converts string to integer value
      if(checkdate(intval($month), intval($day), intval($year)) == 1){
         echo "Valid date input <br>";
         echo $year . "-" . $month . "-" . $day . "<br>";
         $date = $year . "-" . $month . "-" . $day;
         return $date;
      }else{
         die("Not a valid date");
      }
   }
}
?>