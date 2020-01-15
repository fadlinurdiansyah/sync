<?php

/**
 * This class use form send message using telegram @sgeapps_bot
 * make sure you provided token
 * Author @abdmun8
 * 15 Jan 2020
 * 
 * Feature
 * You can use HTML syntax on message
 * You can send to multiple person
 */

header("Access-Control-Allow-Origin: *");

class Telegram
{
    var $token;
    function __construct()
    {
        $this->token = '889173285:AAGWGpoM3owIVvardYu4VYmM9ZpwVr4Roco';
    }

    function sendMessage($chat_id, $text)
    {
        $postdata = http_build_query(
            array(
                'chat_id' => $chat_id,
                'text' => $text,
                'parse_mode' => 'HTML'
            )
        );

        $opts = array(
            'http' =>
            array(
                'method'  => 'POST',
                'header'  => 'Content-Type: application/x-www-form-urlencoded',
                'content' => $postdata
            )
        );

        $context  = stream_context_create($opts);
        $url = "https://api.telegram.org/bot{$this->token}/sendMessage";

        $result = file_get_contents($url, false, $context);
        return $result;
    }
}

// class instance
$telegram = new Telegram();

if (isset($_POST['token']) && $_POST['token'] == '82027888c5bb8fc395411cb6804a066c') {
    $chat_id = $_POST['telegram_id'];
    $text = $_POST['message'];
    $message_sent = 0;
    $notes = "";
    // check if chat id is array
    if (is_array($chat_id)) {
        $length = count($chat_id);
        foreach ($chat_id as $key => $value) {
            $response = $telegram->sendMessage($value, $text);
            $data = json_decode($response);
            if ($data->ok == TRUE) {
                $message_sent++;
            }
        }

        $notes = $message_sent . ' of ' . $length;
    } else {
        $response = $telegram->sendMessage($chat_id, $text);
        $data = json_decode($response);
        if ($data->ok == TRUE) {
            $message_sent++;
        }

        $message_sent . ' of 1';
    }

    $result = [
        'success' => $message_sent > 0 ? TRUE : FALSE,
        'message_sent' => $notes
    ];

    echo json_encode($result);
} else {
    echo json_encode(['success' => FALSE, 'message' => 'Token salah!']);
}
