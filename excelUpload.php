<?php

require('koneksi.php');
require('vendor/autoload.php');


if(isset($_POST['Submit'])){

  $err = "";
  $ekstensi = "";
  $success = "";

  $file_name = $_FILES['filexls']['name'];
  $file_data = $_FILES['filexls']['tmp_name'];

  if (empty($file_name)) {
    $err .= "<li>Masukkan file!</li>";
  } else {
    $ekstensi = pathinfo($file_name)['extension'];
  }

  $ekstensi_allowed = array("xls", "xlsx");
  if (!in_array($ekstensi,$ekstensi_allowed)) {
    $err .= "<li>Masukkan file xls/xlsx!</li>";
  }

  if(empty($err)){
    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($file_data);
    $spreadsheet = $reader->load($file_data);
    $sheetData = $spreadsheet->getActiveSheet()->toArray();

    $jumlahData = 0;

    for ($i=1; $i<count($sheetData); $i++) { 
      $nama_sales = $sheetData[$i]['0'];
      $no_invoice = $sheetData[$i]['1'];
      $date = $sheetData[$i]['2'];
      $cust_id = $sheetData[$i]['3'];
      $customer = $sheetData[$i]['4'];
      $term = $sheetData[$i]['5'];
      $nett = $sheetData[$i]['6'];
      $profit = $sheetData[$i]['7'];
      $urutan = $sheetData[$i]['8'];

      // echo "$nama_sales - $no_invoice - $date - $cust_id - $customer - $term - $nett - $profit\n";
      $query = "insert into penjualan(nama_sales, no_invoice, date, cust_id, customer, term, nett, profit,urutan) values('$nama_sales','$no_invoice','$date','$cust_id','$customer','$term',$nett,$profit,$urutan)";

      mysqli_query($koneksi, $query);

      $jumlahData++;
    }

    if($jumlahData > 0) {
      $success = "Berhasil upload ke database! ".$jumlahData." records!";
    }
  }

  if ($err) {
    ?>
    <div class="alert alert-danger">
      <ul><?php echo $err ?></ul>
    </div>
    <?php
  }

  if($success) {
    ?>
    <div class="alert alert-primary">
      <?php echo $success ?>
    </div>
    <?php
  }


  // $mimes = ['application/vnd.ms-excel','text/xls','text/xlsx','application/vnd.oasis.opendocument.spreadsheet'];
  // if(in_array($_FILES["file"]["type"],$mimes)){


  //   $uploadFilePath = 'uploads/'.basename($_FILES['file']['name']);
  //   move_uploaded_file($_FILES['file']['tmp_name'], $uploadFilePath);


  //   $Reader = new SpreadsheetReader($uploadFilePath);


  //   $totalSheet = count($Reader->sheets());


  //   echo "You have total ".$totalSheet." sheets".


  //   $html="<table border='1'>";
  //   $html.="<tr><th>Title</th><th>Description</th></tr>";


  //   /* For Loop for all sheets */
  //   for($i=0;$i<$totalSheet;$i++){


  //     $Reader->ChangeSheet($i);


  //     foreach ($Reader as $Row)
  //     {
  //       $html.="<tr>";
  //       $nama_sales = isset($Row[0]) ? $Row[0] : '';
  //       $no_invoice = isset($Row[1]) ? $Row[1] : '';
  //       $date = isset($Row[2]) ? $Row[2] : '';
  //       $cust_id = isset($Row[3]) ? $Row[3] : '';
  //       $customer = isset($Row[4]) ? $Row[4] : '';
  //       $term = isset($Row[5]) ? $Row[5] : '';
  //       $nett = isset($Row[6]) ? $Row[6] : '';
  //       $profit = isset($Row[7]) ? $Row[7] : '';

  //       $html.="<td>".$nama_sales."</td>";
  //       $html.="<td>".$no_invoice."</td>";
  //       $html.="<td>".$date."</td>";
  //       $html.="<td>".$cust_id."</td>";
  //       $html.="<td>".$customer."</td>";
  //       $html.="<td>".$term."</td>";
  //       $html.="<td>".$nett."</td>";
  //       $html.="<td>".$profit."</td>";
  //       $html.="</tr>";


  //       $query = "insert into penjualan(nama_sales,no_invoice,date,cust_id,customer,term,nett,profit) values('".$nama_sales."','".$no_invoice."','".$date."','".$cust_id."','".$customer."','".$term."','".$nett."','".$term."','".$nett."','".$profit."')";


  //       $mysqli->query($query);
  //      }


  //   }


  //   $html.="</table>";
  //   echo $html;
  //   echo "<br />Data Inserted in dababase";


  // }else { 
  //   die("<br/>Sorry, File type is not allowed. Only Excel file."); 
  // }


}


?>