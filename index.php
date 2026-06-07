#!/usr/bin/env php
<?php

declare(strict_types=1);

require_once __DIR__ . '/vendor/autoload.php';
date_default_timezone_set('Asia/Kolkata');

// PhpSpreadsheet emits PHP 8.5 deprecation notices from vendor code we cannot
// patch. Hide deprecations so the report stays readable; real warnings and
// errors are still surfaced.
error_reporting(E_ALL & ~E_DEPRECATED);

use Spock\PhpParseMutualFund\ParseCommand;
use Symfony\Component\Console\Application;

$app = new Application();
$app->add(new ParseCommand());
$app->run();
