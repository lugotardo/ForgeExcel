<?php

namespace Lugotardo\Forgeexel\Traits;

use Box\Spout\Writer\Common\Creator\WriterEntityFactory;

/**
 * Trait FormulaTrait
 *
 * Contém todos os métodos relacionados a fórmulas do Excel
 *
 * @package Lugotardo\Forgeexel\Traits
 */
trait FormulaTrait
{
    /**
     * Escreve dados com fórmulas do Excel
     *
     * Exemplo de uso:
     * $dados = [
     *     ['Produto', 'Quantidade', 'Preço', 'Total'],
     *     ['Notebook', 2, 3500, '=B2*C2'],
     *     ['Mouse', 5, 50, '=B3*C3'],
     *     ['Total', '', '', '=SUM(D2:D3)']
     * ];
     *
     * ForgeExcel::writeWithFormulas('vendas.xlsx', $dados);
     *
     * @param string $filePath Caminho do arquivo
     * @param array $data Dados com fórmulas (use string iniciando com =)
     * @param array $headerStyle Estilo opcional para o cabeçalho
     * @return bool TRUE se salvou com sucesso
     */
    public static function writeWithFormulas(
        string $filePath,
        array $data,
        array $headerStyle = [],
    ): bool {
        $directory = dirname($filePath);
        if (!is_dir($directory)) {
            mkdir($directory, 0777, true);
        }

        $writer = WriterEntityFactory::createXLSXWriter();
        $writer->openToFile($filePath);

        // Estilo padrão para header se não especificado
        if (empty($headerStyle) && !empty($data)) {
            $headerStyle = ["bold" => true, "background" => "E0E0E0"];
        }

        foreach ($data as $rowIndex => $rowData) {
            $cells = [];

            foreach ($rowData as $cellValue) {
                $cellStyle = null;

                // Aplica estilo no header (primeira linha)
                if ($rowIndex === 0 && !empty($headerStyle)) {
                    $cellStyle = self::createStyle($headerStyle);
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
}
