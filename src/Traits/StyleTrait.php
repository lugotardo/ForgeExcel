<?php

namespace Lugotardo\Forgeexel\Traits;

use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Common\Entity\Style\Style;
use Box\Spout\Common\Entity\Style\Color;
use Box\Spout\Common\Entity\Style\CellAlignment;
use Box\Spout\Common\Entity\Style\Border;
use Box\Spout\Common\Entity\Style\BorderPart;

/**
 * Trait StyleTrait
 *
 * Contém todos os métodos de formatação e estilos
 *
 * @package Lugotardo\Forgeexel\Traits
 */
trait StyleTrait
{
    /**
     * Cria um estilo personalizado para células
     *
     * Exemplo de uso:
     * $estilo = ForgeExcel::createStyle([
     *     'bold' => true,
     *     'color' => 'FF0000',
     *     'background' => 'FFFF00',
     *     'fontSize' => 14
     * ]);
     *
     * @param array $options Opções de estilo
     * @return Style Objeto de estilo do Spout
     */
    public static function createStyle(array $options = []): Style
    {
        $style = new Style();

        // Negrito
        if (isset($options["bold"]) && $options["bold"]) {
            $style->setFontBold();
        }

        // Itálico
        if (isset($options["italic"]) && $options["italic"]) {
            $style->setFontItalic();
        }

        // Sublinhado
        if (isset($options["underline"]) && $options["underline"]) {
            $style->setFontUnderline();
        }

        // Tamanho da fonte
        if (isset($options["fontSize"])) {
            $style->setFontSize($options["fontSize"]);
        }

        // Cor da fonte (hexadecimal sem #)
        if (isset($options["color"])) {
            $style->setFontColor($options["color"]);
        }

        // Cor de fundo (hexadecimal sem #)
        if (isset($options["background"])) {
            $style->setBackgroundColor($options["background"]);
        }

        // Nome da fonte
        if (isset($options["fontName"])) {
            $style->setFontName($options["fontName"]);
        }

        // Alinhamento horizontal
        if (isset($options["align"])) {
            $style->setCellAlignment($options["align"]);
        }

        // Quebrar texto automaticamente
        if (isset($options["wrapText"]) && $options["wrapText"]) {
            $style->setShouldWrapText(true);
        }

        // Bordas
        if (isset($options["border"])) {
            $borderStyle = $options["borderStyle"] ?? Border::STYLE_THIN;
            $borderColor = $options["borderColor"] ?? Color::BLACK;

            $border = new Border(
                new BorderPart(Border::TOP, $borderColor, $borderStyle),
                new BorderPart(Border::RIGHT, $borderColor, $borderStyle),
                new BorderPart(Border::BOTTOM, $borderColor, $borderStyle),
                new BorderPart(Border::LEFT, $borderColor, $borderStyle),
            );

            $style->setBorder($border);
        }

        return $style;
    }

    /**
     * Escreve dados com formatação personalizada
     *
     * Exemplo de uso:
     * $dados = [
     *     ['Nome', 'Valor', 'Status'],
     *     ['Produto A', 100, 'Ativo'],
     *     ['Produto B', 200, 'Inativo']
     * ];
     *
     * $estilos = [
     *     0 => ['bold' => true, 'background' => 'CCCCCC'], // Header
     *     2 => ['color' => 'FF0000'] // Linha 2 em vermelho
     * ];
     *
     * ForgeExcel::writeWithStyle('arquivo.xlsx', $dados, $estilos);
     *
     * @param string $filePath Caminho do arquivo
     * @param array $data Dados a serem escritos
     * @param array $styles Array associativo [numero_linha => opcoes_estilo]
     * @param array $columnStyles Array associativo [numero_coluna => opcoes_estilo]
     * @return bool TRUE se salvou com sucesso
     */
    public static function writeWithStyle(
        string $filePath,
        array $data,
        array $styles = [],
        array $columnStyles = [],
    ): bool {
        $directory = dirname($filePath);
        if (!is_dir($directory)) {
            mkdir($directory, 0777, true);
        }

        $writer = WriterEntityFactory::createXLSXWriter();
        $writer->openToFile($filePath);

        foreach ($data as $rowIndex => $rowData) {
            $cells = [];

            foreach ($rowData as $colIndex => $cellValue) {
                // Determina o estilo da célula
                $cellStyle = null;

                // Estilo por linha tem prioridade
                if (isset($styles[$rowIndex])) {
                    $cellStyle = self::createStyle($styles[$rowIndex]);
                }
                // Senão, verifica estilo por coluna
                elseif (isset($columnStyles[$colIndex])) {
                    $cellStyle = self::createStyle($columnStyles[$colIndex]);
                }

                $cells[] = WriterEntityFactory::createCell(
                    $cellValue,
                    $cellStyle,
                );
            }

            $row = WriterEntityFactory::createRow($cells);
            $writer->addRow($row);
        }

        $writer->close();
        return true;
    }

    /**
     * Cria uma tabela formatada com estilos predefinidos
     *
     * Exemplo de uso:
     * $dados = [
     *     ['Nome', 'Email', 'Status'],
     *     ['João', 'joao@email.com', 'Ativo'],
     *     ['Maria', 'maria@email.com', 'Inativo']
     * ];
     *
     * ForgeExcel::writeTable('tabela.xlsx', $dados, 'blue');
     *
     * @param string $filePath Caminho do arquivo
     * @param array $data Dados da tabela (primeira linha é header)
     * @param string $theme Tema: 'blue', 'green', 'red', 'orange', 'purple'
     * @return bool TRUE se salvou com sucesso
     */
    public static function writeTable(
        string $filePath,
        array $data,
        string $theme = "blue",
    ): bool {
        $themes = [
            "blue" => [
                "header" => [
                    "bold" => true,
                    "color" => "FFFFFF",
                    "background" => "4472C4",
                ],
                "odd" => ["background" => "D9E1F2"],
                "even" => ["background" => "FFFFFF"],
            ],
            "green" => [
                "header" => [
                    "bold" => true,
                    "color" => "FFFFFF",
                    "background" => "70AD47",
                ],
                "odd" => ["background" => "E2EFDA"],
                "even" => ["background" => "FFFFFF"],
            ],
            "red" => [
                "header" => [
                    "bold" => true,
                    "color" => "FFFFFF",
                    "background" => "C55A11",
                ],
                "odd" => ["background" => "FCE4D6"],
                "even" => ["background" => "FFFFFF"],
            ],
            "orange" => [
                "header" => [
                    "bold" => true,
                    "color" => "FFFFFF",
                    "background" => "ED7D31",
                ],
                "odd" => ["background" => "FBE5D6"],
                "even" => ["background" => "FFFFFF"],
            ],
            "purple" => [
                "header" => [
                    "bold" => true,
                    "color" => "FFFFFF",
                    "background" => "7030A0",
                ],
                "odd" => ["background" => "E4DFEC"],
                "even" => ["background" => "FFFFFF"],
            ],
        ];

        $themeConfig = $themes[$theme] ?? $themes["blue"];

        $directory = dirname($filePath);
        if (!is_dir($directory)) {
            mkdir($directory, 0777, true);
        }

        $writer = WriterEntityFactory::createXLSXWriter();
        $writer->openToFile($filePath);

        foreach ($data as $rowIndex => $rowData) {
            $cells = [];

            // Define o estilo baseado na linha
            $cellStyle = null;
            if ($rowIndex === 0) {
                $cellStyle = self::createStyle($themeConfig["header"]);
            } elseif ($rowIndex % 2 === 1) {
                $cellStyle = self::createStyle($themeConfig["odd"]);
            } else {
                $cellStyle = self::createStyle($themeConfig["even"]);
            }

            foreach ($rowData as $cellValue) {
                $cells[] = WriterEntityFactory::createCell(
                    $cellValue,
                    $cellStyle,
                );
            }

            $row = WriterEntityFactory::createRow($cells);
            $writer->addRow($row);
        }

        $writer->close();
        return true;
    }

    /**
     * Cria múltiplas abas com formatação
     *
     * Exemplo de uso:
     * $abas = [
     *     'Vendas' => [
     *         'data' => [['Produto', 'Valor'], ['Notebook', 3500]],
     *         'headerStyle' => ['bold' => true, 'background' => '4472C4', 'color' => 'FFFFFF']
     *     ],
     *     'Resumo' => [
     *         'data' => [['Total', '=SUM(Vendas!B2:B10)']],
     *         'headerStyle' => ['bold' => true, 'background' => '70AD47', 'color' => 'FFFFFF']
     *     ]
     * ];
     *
     * ForgeExcel::writeStyledSheets('relatorio.xlsx', $abas);
     *
     * @param string $filePath Caminho do arquivo
     * @param array $sheets Array [nome_aba => ['data' => dados, 'headerStyle' => estilo]]
     * @return bool TRUE se salvou com sucesso
     */
    public static function writeStyledSheets(
        string $filePath,
        array $sheets,
    ): bool {
        $directory = dirname($filePath);
        if (!is_dir($directory)) {
            mkdir($directory, 0777, true);
        }

        $writer = WriterEntityFactory::createXLSXWriter();
        $writer->openToFile($filePath);

        $isFirstSheet = true;

        foreach ($sheets as $sheetName => $sheetConfig) {
            if (!$isFirstSheet) {
                $writer->addNewSheetAndMakeItCurrent();
            }

            $currentSheet = $writer->getCurrentSheet();
            $currentSheet->setName($sheetName);

            $data = $sheetConfig["data"] ?? [];
            $headerStyle = $sheetConfig["headerStyle"] ?? null;
            $rowStyles = $sheetConfig["rowStyles"] ?? [];

            foreach ($data as $rowIndex => $rowData) {
                $cells = [];

                $cellStyle = null;
                if ($rowIndex === 0 && $headerStyle !== null) {
                    $cellStyle = self::createStyle($headerStyle);
                } elseif (isset($rowStyles[$rowIndex])) {
                    $cellStyle = self::createStyle($rowStyles[$rowIndex]);
                }

                foreach ($rowData as $cellValue) {
                    $cells[] = WriterEntityFactory::createCell(
                        $cellValue,
                        $cellStyle,
                    );
                }

                $row = WriterEntityFactory::createRow($cells);
                $writer->addRow($row);
            }

            $isFirstSheet = false;
        }

        $writer->close();
        return true;
    }

    /**
     * Estilos de cores predefinidos para facilitar o uso
     *
     * @return array Array com cores comuns
     */
    public static function colors(): array
    {
        return [
            "black" => "000000",
            "white" => "FFFFFF",
            "red" => "FF0000",
            "green" => "00FF00",
            "blue" => "0000FF",
            "yellow" => "FFFF00",
            "orange" => "FFA500",
            "purple" => "800080",
            "pink" => "FFC0CB",
            "gray" => "808080",
            "light_gray" => "D3D3D3",
            "dark_gray" => "A9A9A9",
            "cyan" => "00FFFF",
            "magenta" => "FF00FF",
            "lime" => "00FF00",
            "navy" => "000080",
            "teal" => "008080",
            "olive" => "808000",
            "maroon" => "800000",
            "aqua" => "00FFFF",
        ];
    }

    /**
     * Constantes de alinhamento para facilitar o uso
     *
     * @return array Array com constantes de alinhamento
     */
    public static function alignments(): array
    {
        return [
            "left" => CellAlignment::LEFT,
            "center" => CellAlignment::CENTER,
            "right" => CellAlignment::RIGHT,
        ];
    }
}
