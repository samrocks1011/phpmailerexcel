<?php


require 'vendor/autoload.php';
//require 'C:\wamp64\www\mailer\vendor\phpmailer\phpmailer\src\PHPMailer.php';
//require 'C:\wamp64\www\mailer\vendor\phpmailer\phpmailer\src\SMTP.php';
//require 'C:\wamp64\www\mailer\vendor\phpmailer\phpmailer\src\Exception.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
//use PHPMailer\PHPMailer\PHPMailer;
// use PHPMailer\PHPMailer\SMTP;
//use PHPMailer\PHPMailer\Exception;

function send_SMTP_mail($from,$To,$Cc, $Bcc,$subject,$body,$host,$username,$password,$port)
{
	
	
    // <!-- DO NOT RUN THIS File More Then 2 TIMES ITS A PAID SERVICE -->
    $mail = new PHPMailer;
    $to_array=explode(",",$To);

	
       
        
        $mail->SMTPDebug = 3;                               // Enable verbose debug output

        $mail->isSMTP(); // Set mailer to use SMTP
        $mail->Host = $host; // Specify main and backup SMTP servers
        $mail->SMTPAuth = false; // Enable SMTP authentication
        $mail->Username = $username; // SMTP username
        $mail->Password = $password; // SMTP password
        // $mail->SMTPSecure = ''; // Enable TLS encryption, `ssl` also accepted
        // $mail->Port = 1005; // TCP port to connect to 587,999,2525,1005
        $mail->Port = $port; // TCP port to connect to 587,999,2525,1005
        //$mail->Port = 26; // TCP port to connect to 587,999,2525,1005

        $mail->From = $from;
        $mail->FromName = 'IAT_Test';

        foreach($to_array as $email)
                {
                    $mail->addAddress($email);
                }

        $mail->addCC($Cc);
        $mail->addBCC($Bcc);

        $mail->WordWrap = 50; // Set word wrap to 50 characters
        $mail->addAttachment('test.xlsx');         // Add attachments
        // $mail->addAttachment('/tmp/image.jpg', 'new.jpg');    // Optional name
        $mail->isHTML(false);                                  // Set email format to HTML

        $mail->Subject = $subject;
        $mail->Body = $body;
        $mail->AltBody = $body;

        if (!$mail->send()) {
            echo 'Message could not be sent.';
            echo 'Mailer Error: ' . $mail->ErrorInfo;
        } else {
            echo 'Message has been sent';
        }
   
}



// Main Function of Programee

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
$spreadsheet = $reader->load("test.xlsx");

$d = $spreadsheet->getSheet(0)->toArray();
$sheetData = $spreadsheet->getActiveSheet()->toArray();
$i = 1;
unset($sheetData[0]);

#for Reading To for Mail
foreach ($sheetData as $t) {
    if($t[0]!="" || $t[2]!="" || $t[4]!="" || $t[6]!=""){
        print_r("True");
        // echo ($t[15]);
        print_r($t);
		send_SMTP_mail($t[0],$t[2],$t[4],$t[6],$t[1],$t[7],$t[9],$t[11],$t[13],$t[15]);
    }
   
    $i=$i+1;

}





?>