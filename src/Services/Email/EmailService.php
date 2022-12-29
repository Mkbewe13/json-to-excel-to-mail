<?php

namespace TmeApp\Services\Email;

use Exception;
use Symfony\Component\Mailer\Exception\TransportExceptionInterface;
use Symfony\Component\Mailer\Mailer;
use Symfony\Component\Mailer\Transport;
use Symfony\Component\Mime\Email;

/**
 * Email handling service class.
 */
class EmailService
{
    private Email $emailMessage;
    private Mailer $mailer;
    private Transport\TransportInterface $transport;

    /**
     * Sets all necessary properties for email handling
     *
     * @throws Exception
     */
    public function __construct(Email $emailMessage)
    {
        $this->emailMessage = $emailMessage;
        $this->transport = \Symfony\Component\Mailer\Transport::fromDsn($this->getDSN());
        $this->mailer = new \Symfony\Component\Mailer\Mailer($this->transport);
    }

    /**
     * Return env constant MAILER_DSN, throws exception if constant doesn't exist.
     *
     * @return mixed
     * @throws Exception
     */
    private function getDSN()
    {
        if (!isset($_ENV['MAILER_DSN']) || !$_ENV['MAILER_DSN']) {
            throw new Exception('MAILER_DSN nie jest ustawiony. Sprawdź plik .env');
        }

        return $_ENV['MAILER_DSN'];
    }

    /**
     * Sends email with email object given to class constructor
     *
     * @return void
     * @throws TransportExceptionInterface
     */
    public function sendEmail()
    {
        $this->mailer->send($this->emailMessage);
    }

}
