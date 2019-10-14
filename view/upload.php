<?php

function upload($sftp,$fileName=''){	
    $date = date('Ymd');
    $srcFile = "C:/incoming_aia/incoming_daily/upload/total/$fileName";
    $dstFile = "/FOLDER_AAI/old/test/$fileName";

    $sftpStream = @fopen('ssh2.sftp://'.$sftp.$dstFile, 'w');
    $data_to_send = '';
    try {

        if (!$sftpStream) {
            throw new Exception("Could not open remote file: $dstFile");
        }
       
        $data_to_send = @file_get_contents($srcFile);			
       
        if ($data_to_send === false) {
            throw new Exception("Could not open local file: $srcFile.");
        }
       
        if (@fwrite($sftpStream, $data_to_send) === false) {
            throw new Exception("Could not send data from file: $srcFile.");
        }
       
        fclose($sftpStream);
                       
    } catch (Exception $e) {
        error_log('Exception: ' . $e->getMessage());
        fclose($sftpStream);
    }
return $data_to_send;
}

if (isset($_POST['submit'])){    
    include_once '../config/connection.php';
    $fileName = $_POST['fileToImport'];       
    upload($sftp,$fileName);
}


?>


<!doctype html>
<html lang="en">
  <head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" crossorigin="anonymous">

    <title>AAI Enrollment</title>
  </head>
  <body>
  <div class="container">
    <nav class="navbar navbar-expand-lg navbar-light bg-light">
        <a class="navbar-brand" href="#">AAI International</a>
        <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarTogglerDemo02" aria-controls="navbarTogglerDemo02" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
        </button>
        
            <div class="collapse navbar-collapse" id="navbarTogglerDemo02">
                <ul class="navbar-nav mr-auto mt-2 mt-lg-0">
                <li class="nav-item active">
                    <a class="nav-link" href="#">Home <span class="sr-only">(current)</span></a>
                </li>
                <li class="nav-item dropdown">
                    <a class="nav-link dropdown-toggle" href="#" id="navbarDropdown" role="button" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                        Enrollment
                    </a>
                    <div class="dropdown-menu" aria-labelledby="navbarDropdown">
                        <a class="dropdown-item" href="#">Upload</a>       
                        <a class="dropdown-item" href="import.php">Import</a>             
                        <div class="dropdown-divider"></div>
                        <!-- <a class="dropdown-item" href="#">Something else here</a> -->
                    </div>
                </li>            
                </ul>
                <form class="form-inline my-2 my-lg-0">
                <input class="form-control mr-sm-2" type="search" placeholder="Search">
                <button class="btn btn-outline-success my-2 my-sm-0" type="submit">Search</button>
                </form>
            </div>
        
    </nav>
    </div>
    <div class="container" style="padding-right: 0">
            <div class="float-left" style="width: 47%;">
                <form method="post" action="import.php">
                    <div class="container mt-3 jumbotron">
                        <div style="margin-top: -50px; padding: 10px;" class="bg-info">
                            <h4>Upload File</h4>
                        </div>
                        <hr />                
                        <div class="row">
                            <div class="form-group col-sm-3">
                                <label for="fileToImport">Select File</label>
                            </div>
                            <div class="col-sm-7">
                                <input type="file" class="form-control-file" id="fileToImport" name="fileToImport">
                            </div>
                        </div>
                        <div class="row">
                            <div class="form-group col-sm-3">
                                <label for="fileToImport"></label>
                            </div>
                            <div class="col-sm-3">
                                <input class="btn btn-primary" type="submit" name="submit" value="Submit">
                            </div>
                        </div>

                    </div>
                </form>
            </div>
        </div>


    <!-- Optional JavaScript -->
    <!-- jQuery first, then Popper.js, then Bootstrap JS -->
    <script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" integrity="sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js" integrity="sha384-UO2eT0CpHqdSJQ6hJty5KVphtPhzWj9WO1clHTMGa3JDZwrnQq4sF86dIHNDz0W1" crossorigin="anonymous"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js" integrity="sha384-JjSmVgyd0p3pXB1rRibZUAYoIIy6OrQ6VrjIEaFf/nJGzIxFDsf4x0xIM+B07jRM" crossorigin="anonymous"></script>
  </body>
</html>