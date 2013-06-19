<?php

class StylePhpExcel
{

    protected $style;

    public function setStyle(array $styleAry)
    {
        $this->style = $styleAry;
        return $this;
    }

    public function getStyle()
    {
        return $this->style;
    }

}