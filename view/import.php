<?php
// session_start();
/** 
| Header include for microframework
|========================================================== */
include '../root.php';

use Pusaka\Microframework\Loader;
//===========================================================

Loader::lib('import');
Loader::lib('export');

$root = ROOTDIR;
if (isset($_POST['submit'])) {
    include 'column_mapping.php';
    include $root . 'controller/Enrollment.php';
    $enroll_obj = new Enrollment();
    $fileName = $_POST['fileToImport'];
    $client = $_POST['client'];
    $company = $_POST['company'];
    $enroll_obj->insert($fileName, $column, $client, $company);
}
if (isset($_POST['submit_principle'])) {
    include $root . 'controller/Enrollment.php';
    $enroll_obj = new Enrollment();
    $client = $_POST['client'];
    $company = $_POST['company'];
    $enroll_obj->update($client, $company);
}
if (isset($_POST['submitLapse'])) {
    include 'column_mapping.php';
    include $root . 'controller/Enrollment.php';
    $fileName = $_POST['fileToImport'];
    $enroll_obj = new Enrollment;
    $res = $enroll_obj->updateLapse($fileName, $column);
    $_SESSION['lapse_succed'] = $res;
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
    <script src="../framework/assets/js/jquery-3.4.1.slim.min.js"></script>
    <title>AAI Enrollment</title>
</head>

<body>
    <div class="container bg-secondary">
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
                            <a class="dropdown-item" href="upload.php">Upload</a>
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
    <div class="container">
        <div class="container mt-2">
            <?php
            if (isset($_POST['submitLapse']) or isset($_POST['submit'])) {
                if ($_SESSION['lapse_succed'] == 1) {
                    echo '<div class="mx-auto" style="width: 80%">
                <div class="alert alert-success" role="alert">
                    Success!
                </div>
            </div>'
                    ?>
            <?php
                }
            }
            ?>

            <div class="row">
                <div class="col align-self-start">

                </div>
                <div class="col align-self-center">
                    <div class="container bg-info" >
                        <div class="form-group">
                            <label for="exampleFormControlSelect1">Action</label>
                            <select class="form-control" id="actionSelect" onchange="actionSelect()">
                                <option value="1">1. New Member (Card Print)</option>
                                <option value="2">2. New Member (No Card Print)</option>
                                <option value="3">3. Lapse</option>
                                <option value="4">4. Reinstate</option>
                                <option value="5">5. Terminate</option>
                                <option value="6">6. Suspend</option>
                                <option value="7">7. Unsuspend</option>
                                <option value="8">8. Update Plan ( With card print)</option>
                                <option value="9">9. Update Plan ( Without card print)</option>
                                <option value="10">10. Update Data ( With card print )</option>
                                <option value="11">11. Update Data ( Without card print)</option>
                                <option value="12">12. Renewal ( With card print)</option>
                                <option value="13">13. Renewal ( Without card print)</option>
                                <option value="14">14. Reprint Card</option>
                            </select>
                        </div>
                    </div>
                </div>
                <div class="col align-self-end">

                </div>
            </div>
        </div>
        <div class="container" style="width: 100%">
            <div id="one" class="container" style="padding-right: 0">
                <div class="mx-auto" style="width: 80%;">
                    <form method="post" action="import.php">
                        <div class="container mt-3 jumbotron">
                            <div style="margin-top: -50px; padding: 10px; background-color: #75C0E0;" >
                                <h4>Import Enrollment</h4>
                            </div>
                            <hr />
                            <div class="row">
                                <div class="form-group col-sm-3">
                                    <label for="client">Client</label>
                                </div>
                                <div class="col-sm-6">
                                    <input type="text" class="form-control" required id="client" name="client">
                                </div>
                            </div>
                            <div class="row">
                                <div class="form-group col-sm-3">
                                    <label for="company">Company</label>
                                </div>
                                <div class="col-sm-6">
                                    <input type="text" class="form-control" required id="company" name="company">
                                </div>
                            </div>
                            <div class="row">
                                <div class="form-group col-sm-3">
                                    <label for="fileToImport">Input File</label>
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
            <div id="three" class="container d-none" style="padding-left: 0">
                <div class="mx-auto" style="width: 80%">
                    <form method="post" action="import.php">
                        <div class="container mt-3 jumbotron">
                            <div style="margin-top: -50px; padding: 10px;" class="bg-warning">
                                <h4>Lapse</h4>
                            </div>
                            <hr />
                            <div class="row">
                                <div class="form-group col-sm-3">
                                    <label for="fileToImport">Input File</label>
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
                                    <input class="btn btn-primary" type="submit" name="submitLapse" value="Submit">
                                </div>
                            </div>
                        </div>
                    </form>
                </div>
            </div>
        </div>

        <script>
            // function actionSelect() {
            //     var newE = document.getElementById('actionSelect').value;
            //     if (newE != 1) {
            //         document.getElementById('1').style.display = 'none';
            //     } else {
            //         document.getElementById('1').style.display = 'block';
            //     }
            // }

            $('#actionSelect').change(function() {
                var newE = $(this).val();
                if (newE > 2) {
                    $('#one').hide();
                } else {
                    $('#one').show();
                }
            })

            $('#actionSelect').change(function() {
                var newE = $(this).val();
                if (newE != 3) {
                    $('#three').attr('class', "container d-none");
                } else {
                    $('#three').attr('class', "container");
                }
            })
        </script>

        <!-- Optional JavaScript -->
        <!-- jQuery first, then Popper.js, then Bootstrap JS -->
        <script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" integrity="sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo" crossorigin="anonymous"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js" integrity="sha384-UO2eT0CpHqdSJQ6hJty5KVphtPhzWj9WO1clHTMGa3JDZwrnQq4sF86dIHNDz0W1" crossorigin="anonymous"></script>
        <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js" integrity="sha384-JjSmVgyd0p3pXB1rRibZUAYoIIy6OrQ6VrjIEaFf/nJGzIxFDsf4x0xIM+B07jRM" crossorigin="anonymous"></script>
</body>

</html>