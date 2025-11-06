<?php

namespace Lugotardo\Forgeexel;

use Lugotardo\Forgeexel\Traits\ReadTrait;
use Lugotardo\Forgeexel\Traits\WriteTrait;
use Lugotardo\Forgeexel\Traits\StyleTrait;
use Lugotardo\Forgeexel\Traits\FormulaTrait;
use Lugotardo\Forgeexel\Traits\UtilityTrait;

/**
 * ForgeExcel - Classe simplificada para manipulação de arquivos Excel
 *
 * Esta classe facilita a leitura e escrita de arquivos Excel (XLSX, CSV, ODS)
 * sem complicação. Use métodos simples e diretos!
 *
 * @author Luan Gotardo <luan.gotardo.dev@gmail.com>
 * @package Lugotardo\Forgeexel
 */
class ForgeExcel
{
    // Importa métodos de leitura de arquivos
    use ReadTrait;

    // Importa métodos de escrita básica
    use WriteTrait;

    // Importa métodos de formatação e estilos
    use StyleTrait;

    // Importa métodos de fórmulas
    use FormulaTrait;

    // Importa métodos utilitários
    use UtilityTrait;
}
