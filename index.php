#!/usr/bin/env php
<?php
require_once __DIR__ . '/vendor/autoload.php';
date_default_timezone_set('Asia/Kolkata');

use Spock\PhpParseMutualFund\ParseCommand;
use Symfony\Component\Console\Application;

$app = new Application();
$app->add(new ParseCommand());
$app->run();
