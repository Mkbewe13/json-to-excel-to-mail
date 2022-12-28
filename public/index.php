<?php

use Symfony\Component\Mime\Email;

require_once __DIR__.'/../vendor/autoload.php';
$dotenv = Dotenv\Dotenv::createImmutable(__DIR__.'/../');
$dotenv->load();
$loader = new \Twig\Loader\FilesystemLoader(__DIR__.'/../templates');
$twig = new \Twig\Environment($loader);

$test = new \TmeApp\Services\Xlsx\SpreadsheetService();
$test->getXlsx();



$email = (new Email())->from('tmeapp@example.com')->to('someone@example.com')->subject('Products sheet')->text('Products data xlsx file')->attachFromPath('demo.xlsx');
$emailService = new \TmeApp\Services\Email\EmailService($email);
$emailService->sendEmail();

echo $twig->render('home.twig');
