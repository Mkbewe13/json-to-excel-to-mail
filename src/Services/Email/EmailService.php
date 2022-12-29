<?php

namespace TmeApp\Services\Email;

use Symfony\Component\Mailer\Mailer;
use Symfony\Component\Mailer\Transport;
use Symfony\Component\Mime\Email;

class EmailService
{
    private Email $emailMessage;
    private Mailer $mailer;
    private Transport\TransportInterface $transport;

    public function __construct(Email $emailMessage)
    {
        $this->emailMessage = $emailMessage;
        $this->transport = \Symfony\Component\Mailer\Transport::fromDsn($this->getDSN());
        $this->mailer = new \Symfony\Component\Mailer\Mailer($this->transport);
    }

    private function getDSN(){
        if(!isset($_ENV['MAILER_DSN']) || !$_ENV['MAILER_DSN']){
            throw new \Exception('MAILER_DSN nie jest ustawiony. Sprawdź plik .env');
        }

        return $_ENV['MAILER_DSN'];
    }

    public function sendEmail(){
        $this->mailer->send($this->emailMessage);
    }

}
