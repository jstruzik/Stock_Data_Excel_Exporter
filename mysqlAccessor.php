<?php
class MySqlDAL
{
   public $connection = null; // hold the connection link resource
   public $query_count = 0; // total number of queries executed

   // Connect when creating an instance of this class
   public function __construct($host, $user, $pass, $db)
   {
      $this->connection = mysqli_connect($host, $user, $pass, $db);
      if (mysqli_connect_errno()) {
         printf("Connect failed: %s\n", mysqli_connect_error());
         exit();
      }
      else{
         echo "Connecting to Database successful.";
      }
      return true;
   }

   // Run a query against the database. The key here is to make sure
   // every query command goes through this function
   public function Query($sql)
   {
      // Execute Query & Get Result Resource
      $result = mysqli_query($this->connection,$sql);

      // Increment Query Counter
      $this->query_count++;

      // Return the result
      return $result;
   }

   // Function to take an SQL query, execute it, and return all
   // the rows in an assoc. array.
   public function FetchArray($sql)
   {
      // Create Empty Array to Store all the rows
      $array = array();
      
      // Execute the Query and get the Result
      if($result = $this->Query($sql)){
         // Loop through each row
         while($row = mysqli_fetch_assoc($result))
         {
            // Add the row to the array
            $array[] = $row;
         }
         
         mysqli_free_result($result);
      }

      // Return the array containing all the rows
      return $array;
   }
}
?>