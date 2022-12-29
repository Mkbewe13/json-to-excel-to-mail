<?php

use Symfony\Component\Mime\Email;

require_once __DIR__ . '/../vendor/autoload.php';

$dotenv = Dotenv\Dotenv::createImmutable(__DIR__ . '/../');
$dotenv->load();

$loader = new \Twig\Loader\FilesystemLoader(__DIR__ . '/../templates');
$twig = new \Twig\Environment($loader);

$message = null;
$success = true;
if (isset($_POST['email']) && $_POST['email']) {
    try {
        $spreadSheet = new \TmeApp\Services\Xlsx\SpreadsheetService();
        $spreadSheet->getXlsx();

        $email = (new Email())
            ->from('tmeapp@example.com')
            ->to($_POST['email'])
            ->subject('Dane produktów')
            ->text('W załączniku znajdziesz plik .xlsx z danymi produktów.')
            ->attachFromPath(__DIR__ . '/../var/tmp/products_data.xlsx');
        $emailService = new \TmeApp\Services\Email\EmailService($email);
        $emailService->sendEmail();

        unlink(__DIR__ . '/../var/tmp/products_data.xlsx');
        $message = 'Email z danymi został wysłany na adres: ' . $_POST['email'];

    } catch (Exception $e) {
        $message = $e->getMessage();
        $success = false;
    }
}


try {
    echo $twig->render('home.twig', ['message' => $message,'success' => $success]);
} catch (\Twig\Error\LoaderError|\Twig\Error\RuntimeError|\Twig\Error\SyntaxError $e) {
    printf('Wystąpił błąd krytyczny. Nie udało się wczytać strony: %s',$e->getMessage());
}


