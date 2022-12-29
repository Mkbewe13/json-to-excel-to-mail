<?php
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
        $spreadSheet->createXlsx();

        $emailService = new \TmeApp\Services\Email\EmailService($_POST['email']);
        $emailService->sendEmail();

        $message = 'Email z danymi został wysłany na adres: ' . $_POST['email'];

    } catch (Exception|\Symfony\Component\Mailer\Exception\TransportExceptionInterface $e) {
        $message = $e->getMessage();
        $success = false;
    }
}


try {
    echo $twig->render('home.twig', ['message' => $message, 'success' => $success]);
} catch (\Twig\Error\LoaderError|\Twig\Error\RuntimeError|\Twig\Error\SyntaxError $e) {
    printf('Wystąpił błąd krytyczny. Nie udało się wczytać strony: %s', $e->getMessage());
}



