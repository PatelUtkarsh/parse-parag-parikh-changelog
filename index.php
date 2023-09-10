<?php
// !/usr/bin/env php
// Write a symphony console boilerplate code here.
require_once __DIR__ . '/vendor/autoload.php';

use Symfony\Component\Console\Application;
$app = new Application();
$app->add(new \Spock\PhpParseMutualFund\ParseSheet());
$app -> run();