<?php
 
require_once('mysqlAccessor.php');
 
// Handles the creation and management of MySQL Connections
class MySqlManager
{
 
    // Holds the singleton instance
    private static $instance = null;
 
    // Contians each MySqlDAL object created;
    private $connections;
 
    private function __construct()
    {
        // Its private to prevent creation outside of the GetInstance function
        $this->connections = array();
    }
 
    public static function GetInstance()
    {
        // If the instance is null, make one
        if(!self::$instance)
        {
            self::$instance = new MySqlManager();
        }
 
        return self::$instance;
    }
 
    // Connect to a new database;
    public static function CreateConnection($host, $user, $pass, $name = 'default')
    {
        $manager = self::GetInstance();
        $manager->connections[$name] = new MySqlDAL($host, $user, $pass, $name);
 
        return $manager->connections[$name];
    }
 
    public static function GetConnection($name = 'default')
    {
        $manager = self::GetInstance();
        if(isset($manager->connections[$name]))
        {
            return $manager->connections[$name];
        }
        else
        {
            // handle connection not found error...
        }
    }
}
?>