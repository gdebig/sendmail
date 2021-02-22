<?php
//Import PHPMailer classes into the global namespace
//These must be at the top of your script, not inside a function

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PHPMailer\PHPMailer\SMTP;

require 'phpmailer\src\PHPMailer.php';
require 'phpmailer\src\SMTP.php';
require 'phpmailer\src\Exception.php';

require_once __DIR__.'/simplexlsx/src/SimpleXLSX.php';

$mail = new PHPMailer();

if ( $xlsx = SimpleXLSX::parse('daftar_email_1.xlsx') ) {	
	foreach ( $xlsx->rows() as $k => $r ){
        echo $r[0];echo "<br />";
        echo $r[1];echo "<br />";
        echo $r[2];echo "<br />";
        try {
            //Server settings
            //$mail->SMTPDebug = SMTP::DEBUG_SERVER;                      //Enable verbose debug output
            $mail->isSMTP();                                            //Send using SMTP
            $mail->Host = 'tls://smtp.gmail.com:587';
            $mail->SMTPOptions = array(
                'ssl' => array(
                'verify_peer' => false,
                'verify_peer_name' => false,
                'allow_self_signed' => true
                )
            );                    //Set the SMTP server to send through
            $mail->SMTPAuth   = true;                                   //Enable SMTP authentication
            $mail->Username   = 'sekretariat.ichi@gmail.com';                     //SMTP username
            $mail->Password   = '123Ichi456';                               //SMTP password
            //$mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;         //Enable TLS encryption; `PHPMailer::ENCRYPTION_SMTPS` encouraged
            //$mail->Port       = 587;                                    //TCP port to connect to, use 465 for `PHPMailer::ENCRYPTION_SMTPS` above
        
            //Recipients
            //$mail->setFrom('sekretariat.ichi@gmail.com', 'Sekretariat ICHI');
            
            $mail->clearAllRecipients();
            $mail->clearAttachments();
            $mail->From = 'sekretariat.ichi@gmail.com';
            $mail->FromName = 'Sekretariat ICHI';
            $mail->addAddress($r[0], $r[1]);     //Add a recipient
            //$mail->addAddress('ellen@example.com');               //Name is optional
            $mail->addReplyTo('sekretariat.ichi@gmail.com', 'Sekretariat ICHI');
            //$mail->addCC('cc@example.com');
            //$mail->addBCC('bcc@example.com');
        
            //Attachments
            //$mail->addAttachment('/var/tmp/file.tar.gz');         //Add attachments
            //$mail->addAttachment('/tmp/image.jpg', 'new.jpg');    //Optional name
        
            //Content
            $mail->isHTML(true);                                  //Set email format to HTML
            $mail->Subject = 'This is test email3';
            $mail->Body    = 'This is the HTML message body <b>in bold!</b>';
            $mail->AltBody = 'This is the body in plain text for non-HTML mail clients';
            $mail->AddAttachment('attachments\\'.$r[2].'.pdf');
        
            $mail->send();
            echo 'Message has been sent <br />';
        } catch (Exception $e) {
            echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
        }
    }
} else {
	echo SimpleXLSX::parseError();
}
?>
