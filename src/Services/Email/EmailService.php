<?php

namespace TmeApp\Services\Email;

use Exception;
use Symfony\Component\Mailer\Exception\TransportExceptionInterface;
use Symfony\Component\Mailer\Mailer;
use Symfony\Component\Mailer\Transport;
use Symfony\Component\Mime\Email;
use TmeApp\Services\Xlsx\SpreadsheetService;

/**
 * Email handling service class.
 */
class EmailService
{
    private Email $emailMessage;
    private Mailer $mailer;
    private Transport\TransportInterface $transport;
    private ?string $customAttachmentPath = null;

    private const DEFAULT_EMAIL_FROM = 'tmeapp@example.com';
    private const DEFAULT_EMAIL_SUBJECT = 'Dane produktów';
    private const DEFAULT_EMAIL_TEXT = 'W załączniku znajdziesz plik .xlsx z danymi produktów.';

    /**
     * Sets all necessary properties for email handling
     *
     * @throws Exception
     */
    public function __construct(string $recipient)
    {
        $this->transport = \Symfony\Component\Mailer\Transport::fromDsn($this->getDSN());
        $this->mailer = new \Symfony\Component\Mailer\Mailer($this->transport);
        $this->prepareEmail($recipient);
        $this->attachProductDataFile();
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

        $filePath = $this->customAttachmentPath ?? SpreadsheetService::getDefaultXlsxFilePath();

        if(file_exists($filePath)){
            unlink($filePath);
        }
    }

    /**
     * @param string $recipient
     * @param string $from
     * @param string $subject
     * @param string $content
     * @return void
     */
    public function prepareEmail(string $recipient,string $from = self::DEFAULT_EMAIL_FROM, string $subject = self::DEFAULT_EMAIL_SUBJECT,string $content = self::DEFAULT_EMAIL_TEXT){
        $this->emailMessage = new Email();
        $this->emailMessage->to($recipient);
        $this->emailMessage->from($from);
        $this->emailMessage->subject($subject);
        $this->emailMessage->text($content);
    }

    /**
     * @throws Exception
     */
    public function attachProductDataFile(string $path = null){

        $this->customAttachmentPath = $path;
        $path = $this->customAttachmentPath ?? SpreadsheetService::getDefaultXlsxFilePath();

        if(!file_exists($path)){
            throw new Exception('Plik z danymi nie istnieje.');
        }

        try {
            $this->emailMessage->attachFromPath($path);
        }catch (Exception $e){
            throw new Exception('Wystąpił błąd podczas dodawania załącznika.');
        }

    }

}
