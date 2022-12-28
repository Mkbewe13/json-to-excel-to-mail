<?php

require_once __DIR__.'/../vendor/autoload.php';

$loader = new \Twig\Loader\FilesystemLoader(__DIR__.'/../templates');
$twig = new \Twig\Environment($loader);

$test = new \TmeApp\Services\XlsxService();
$test->getXlsx();
echo $twig->render('home.twig');
