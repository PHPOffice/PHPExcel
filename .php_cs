<?php

$finder = Symfony\CS\Finder\DefaultFinder::create();

$config = Symfony\CS\Config\Config::create()
    ->finder($finder)
    ->fixers(array(
        'indentation',
        'linefeed',
        'trailing_spaces',
        'unused_use',
        'extra_empty_lines',
        'include',
        'short_tag',
        'php_closing_tag',
        'return',
        'braces',
        'phpdoc_params',
        'controls_spaces',
        'elseif',
        'eof_ending',
    ))
;

return $config;
