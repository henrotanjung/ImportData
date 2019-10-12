
<?php
$port = "22";

$server = "103.41.246.140";
$username = "user_aai1";
$password = "T4usY7z`5s";

$connection = ssh2_connect($server, 22);
if(!ssh2_auth_password($connection, $username, $password))
{
	throw new Exception('error 1');
}	

$sftp = ssh2_sftp($connection);
if(!$sftp)
{
	throw new Exception('error 2');
}

//$sftp = intval($sftp);

// $sftp_fd = intval($sftp);

// $dest_dir = 'FOLDER_AAI/old/test';
// $handle = opendir("ssh2.sftp://$sftp_fd/./FOLDER_AAI/old/test");