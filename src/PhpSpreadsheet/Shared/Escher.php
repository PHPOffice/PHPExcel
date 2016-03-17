<?php

namespace PhpOffice\PhpExcel\Shared;

/**
 * PhpOffice\PhpExcel\Shared\Escher
 *
 * Copyright (c) 2006 - 2016 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PhpOffice\PhpExcel\Shared\Escher
 * @copyright  Copyright (c) 2006 - 2016 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Escher
{
    /**
     * Drawing Group Container
     *
     * @var Escher\DggContainer
     */
    private $dggContainer;

    /**
     * Drawing Container
     *
     * @var Escher\DgContainer
     */
    private $dgContainer;

    /**
     * Get Drawing Group Container
     *
     * @return Escher\DgContainer
     */
    public function getDggContainer()
    {
        return $this->dggContainer;
    }

    /**
     * Set Drawing Group Container
     *
     * @param Escher\DggContainer $dggContainer
     */
    public function setDggContainer($dggContainer)
    {
        return $this->dggContainer = $dggContainer;
    }

    /**
     * Get Drawing Container
     *
     * @return Escher\DgContainer
     */
    public function getDgContainer()
    {
        return $this->dgContainer;
    }

    /**
     * Set Drawing Container
     *
     * @param Escher\DgContainer $dgContainer
     */
    public function setDgContainer($dgContainer)
    {
        return $this->dgContainer = $dgContainer;
    }
}
