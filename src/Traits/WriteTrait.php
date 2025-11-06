<?php

namespace Lugotardo\Forgeexel\Traits;

use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Exception;

/**
 * Trait WriteTrait
 *
 * Contém todos os métodos básicos de escrita de arquivos Excel
 *
 * @package Lugotardo\Forgeexel\Traits
 */
trait WriteTrait
{
    /**
     * Escreve dados em um arquivo Excel
     *
     * Exemplo de uso:
     * $dados = [
     *     ['Nome', 'Email', 'Idade'],
     *     ['João', 'joao@email.com', 25],
     *     ['Maria', 'maria@email.com', 30]
     * ];
     * ForgeExcel::write('saida.xlsx', $dados);
     *
     * @param string $filePath Caminho onde o arquivo será salvo
     * @param array $data Array de dados (cada item é uma linha)
     * @param string $type Tipo do arquivo: 'xlsx', 'csv' ou 'ods'
     * @return bool TRUE se salvou com sucesso
     * @throws Exception Se houver erro ao criar ou escrever o arquivo
     */
    public static function write(
        string $filePath,
        array $data,
        string $type = "xlsx",
    ): bool {
        // Verifica se o diretório existe, senão cria
        $directory = dirname($filePath);
        if (!is_dir($directory)) {
            mkdir($directory, 0777, true);
        }

        // Cria o writer baseado no tipo especificado
        switch (strtolower($type)) {
            case "csv":
                $writer = WriterEntityFactory::createCSVWriter();
                break;
            case "ods":
                $writer = WriterEntityFactory::createODSWriter();
                break;
            case "xlsx":
            default:
                $writer = WriterEntityFactory::createXLSXWriter();
                break;
        }

        // Abre o arquivo para escrita
        $writer->openToFile($filePath);

        // Escreve cada linha de dados
        foreach ($data as $row) {
            // Cria uma linha do Spout
            $rowFromValues = WriterEntityFactory::createRowFromArray($row);
            // Adiciona a linha ao arquivo
            $writer->addRow($rowFromValues);
        }

        // Fecha e salva o arquivo
        $writer->close();

        return true;
    }

    /**
     * Escreve dados em múltiplas abas de um arquivo Excel
     *
     * Exemplo de uso:
     * $abas = [
     *     'Clientes' => [
     *         ['Nome', 'Email'],
     *         ['João', 'joao@email.com']
     *     ],
     *     'Produtos' => [
     *         ['Produto', 'Preço'],
     *         ['Notebook', 3000]
     *     ]
     * ];
     * ForgeExcel::writeWithSheets('saida.xlsx', $abas);
     *
     * @param string $filePath Caminho onde o arquivo será salvo
     * @param array $sheets Array associativo [nome_aba => dados]
     * @return bool TRUE se salvou com sucesso
     * @throws Exception Se houver erro ao criar ou escrever o arquivo
     */
    public static function writeWithSheets(
        string $filePath,
        array $sheets,
    ): bool {
        // Verifica se o diretório existe, senão cria
        $directory = dirname($filePath);
        if (!is_dir($directory)) {
            mkdir($directory, 0777, true);
        }

        // Cria o writer (apenas XLSX e ODS suportam múltiplas abas)
        $extension = pathinfo($filePath, PATHINFO_EXTENSION);
        if (strtolower($extension) === "ods") {
            $writer = WriterEntityFactory::createODSWriter();
        } else {
            $writer = WriterEntityFactory::createXLSXWriter();
        }

        // Abre o arquivo para escrita
        $writer->openToFile($filePath);

        $isFirstSheet = true;

        // Percorre cada aba
        foreach ($sheets as $sheetName => $data) {
            // Se não é a primeira aba, cria uma nova aba
            if (!$isFirstSheet) {
                $writer->addNewSheetAndMakeItCurrent();
            }

            // Define o nome da aba atual
            $currentSheet = $writer->getCurrentSheet();
            $currentSheet->setName($sheetName);

            // Escreve os dados na aba
            foreach ($data as $row) {
                $rowFromValues = WriterEntityFactory::createRowFromArray($row);
                $writer->addRow($rowFromValues);
            }

            $isFirstSheet = false;
        }

        // Fecha e salva o arquivo
        $writer->close();

        return true;
    }
}
