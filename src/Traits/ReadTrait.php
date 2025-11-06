<?php

namespace Lugotardo\Forgeexel\Traits;

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Exception;

/**
 * Trait ReadTrait
 *
 * Contém todos os métodos de leitura de arquivos Excel
 *
 * @package Lugotardo\Forgeexel\Traits
 */
trait ReadTrait
{
    /**
     * Lê um arquivo Excel e retorna todos os dados em array
     *
     * Exemplo de uso:
     * $dados = ForgeExcel::read('planilha.xlsx');
     *
     * @param string $filePath Caminho completo do arquivo Excel
     * @param bool $firstRowAsHeader Se TRUE, usa primeira linha como chave do array
     * @return array Dados da planilha em formato de array
     * @throws Exception Se o arquivo não existir ou houver erro na leitura
     */
    public static function read(
        string $filePath,
        bool $firstRowAsHeader = false,
    ): array {
        // Verifica se o arquivo existe
        if (!file_exists($filePath)) {
            throw new Exception("Arquivo não encontrado: {$filePath}");
        }

        // Cria o leitor baseado na extensão do arquivo
        $reader = ReaderEntityFactory::createReaderFromFile($filePath);

        // Abre o arquivo para leitura
        $reader->open($filePath);

        $data = [];
        $headers = [];
        $isFirstRow = true;

        // Percorre todas as abas da planilha
        foreach ($reader->getSheetIterator() as $sheet) {
            // Percorre todas as linhas da aba atual
            foreach ($sheet->getRowIterator() as $row) {
                // Converte a linha em array
                $rowArray = $row->toArray();

                // Se a primeira linha é header, guarda os nomes das colunas
                if ($firstRowAsHeader && $isFirstRow) {
                    $headers = $rowArray;
                    $isFirstRow = false;
                    continue;
                }

                // Se tem headers, cria array associativo
                if (!empty($headers)) {
                    $rowWithKeys = [];
                    foreach ($rowArray as $index => $value) {
                        $key = $headers[$index] ?? $index;
                        $rowWithKeys[$key] = $value;
                    }
                    $data[] = $rowWithKeys;
                } else {
                    // Senão, adiciona array numérico simples
                    $data[] = $rowArray;
                }
            }
        }

        // Fecha o arquivo
        $reader->close();

        return $data;
    }

    /**
     * Lê apenas a primeira aba de um arquivo Excel
     * Mais rápido quando você só precisa dos dados da primeira aba
     *
     * @param string $filePath Caminho do arquivo
     * @param bool $firstRowAsHeader Se TRUE, usa primeira linha como chave
     * @return array Dados da primeira aba
     * @throws Exception Se o arquivo não existir ou houver erro
     */
    public static function readFirstSheet(
        string $filePath,
        bool $firstRowAsHeader = false,
    ): array {
        if (!file_exists($filePath)) {
            throw new Exception("Arquivo não encontrado: {$filePath}");
        }

        $reader = ReaderEntityFactory::createReaderFromFile($filePath);
        $reader->open($filePath);

        $data = [];
        $headers = [];
        $isFirstRow = true;

        // Pega apenas a primeira aba
        foreach ($reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                $rowArray = $row->toArray();

                if ($firstRowAsHeader && $isFirstRow) {
                    $headers = $rowArray;
                    $isFirstRow = false;
                    continue;
                }

                if (!empty($headers)) {
                    $rowWithKeys = [];
                    foreach ($rowArray as $index => $value) {
                        $key = $headers[$index] ?? $index;
                        $rowWithKeys[$key] = $value;
                    }
                    $data[] = $rowWithKeys;
                } else {
                    $data[] = $rowArray;
                }
            }
            // Para após a primeira aba
            break;
        }

        $reader->close();

        return $data;
    }

    /**
     * Lê todas as abas de um arquivo Excel separadamente
     * Retorna um array associativo com o nome de cada aba e seus dados
     *
     * Exemplo de uso:
     * $todasAbas = ForgeExcel::readAllSheets('arquivo.xlsx', true);
     * foreach ($todasAbas as $nomeAba => $dados) {
     *     echo "Aba: {$nomeAba} tem " . count($dados) . " linhas\n";
     * }
     *
     * @param string $filePath Caminho do arquivo
     * @param bool $firstRowAsHeader Se TRUE, usa primeira linha como chave
     * @return array Array associativo [nome_aba => dados]
     * @throws Exception Se o arquivo não existir
     */
    public static function readAllSheets(
        string $filePath,
        bool $firstRowAsHeader = false,
    ): array {
        if (!file_exists($filePath)) {
            throw new Exception("Arquivo não encontrado: {$filePath}");
        }

        $reader = ReaderEntityFactory::createReaderFromFile($filePath);
        $reader->open($filePath);

        $allSheets = [];

        // Percorre todas as abas
        foreach ($reader->getSheetIterator() as $sheet) {
            $sheetName = $sheet->getName();
            $sheetData = [];
            $headers = [];
            $isFirstRow = true;

            foreach ($sheet->getRowIterator() as $row) {
                $rowArray = $row->toArray();

                if ($firstRowAsHeader && $isFirstRow) {
                    $headers = $rowArray;
                    $isFirstRow = false;
                    continue;
                }

                if (!empty($headers)) {
                    $rowWithKeys = [];
                    foreach ($rowArray as $index => $value) {
                        $key = $headers[$index] ?? $index;
                        $rowWithKeys[$key] = $value;
                    }
                    $sheetData[] = $rowWithKeys;
                } else {
                    $sheetData[] = $rowArray;
                }
            }

            $allSheets[$sheetName] = $sheetData;
        }

        $reader->close();

        return $allSheets;
    }

    /**
     * Conta quantas linhas tem um arquivo Excel
     * Útil para arquivos grandes antes de processar
     *
     * @param string $filePath Caminho do arquivo
     * @param bool $countHeader Se FALSE, não conta a primeira linha
     * @return int Número de linhas
     * @throws Exception Se o arquivo não existir
     */
    public static function countRows(
        string $filePath,
        bool $countHeader = true,
    ): int {
        if (!file_exists($filePath)) {
            throw new Exception("Arquivo não encontrado: {$filePath}");
        }

        $reader = ReaderEntityFactory::createReaderFromFile($filePath);
        $reader->open($filePath);

        $totalRows = 0;
        $isFirstRow = true;

        foreach ($reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                if (!$countHeader && $isFirstRow) {
                    $isFirstRow = false;
                    continue;
                }
                $totalRows++;
            }
        }

        $reader->close();

        return $totalRows;
    }

    /**
     * Lê um arquivo Excel em lotes (chunks)
     * Perfeito para processar arquivos muito grandes sem estourar a memória
     *
     * Exemplo de uso:
     * ForgeExcel::readInChunks('grande.xlsx', 100, function($lote) {
     *     foreach ($lote as $linha) {
     *         // Processa cada linha
     *         echo $linha[0] . "\n";
     *     }
     * });
     *
     * @param string $filePath Caminho do arquivo
     * @param int $chunkSize Quantas linhas processar por vez
     * @param callable $callback Função que recebe cada lote de dados
     * @param bool $firstRowAsHeader Se TRUE, usa primeira linha como chave
     * @return void
     * @throws Exception Se o arquivo não existir
     */
    public static function readInChunks(
        string $filePath,
        int $chunkSize,
        callable $callback,
        bool $firstRowAsHeader = false,
    ): void {
        if (!file_exists($filePath)) {
            throw new Exception("Arquivo não encontrado: {$filePath}");
        }

        $reader = ReaderEntityFactory::createReaderFromFile($filePath);
        $reader->open($filePath);

        $currentChunk = [];
        $headers = [];
        $isFirstRow = true;

        foreach ($reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                $rowArray = $row->toArray();

                // Captura o header se necessário
                if ($firstRowAsHeader && $isFirstRow) {
                    $headers = $rowArray;
                    $isFirstRow = false;
                    continue;
                }

                // Monta a linha com ou sem headers
                if (!empty($headers)) {
                    $rowWithKeys = [];
                    foreach ($rowArray as $index => $value) {
                        $key = $headers[$index] ?? $index;
                        $rowWithKeys[$key] = $value;
                    }
                    $currentChunk[] = $rowWithKeys;
                } else {
                    $currentChunk[] = $rowArray;
                }

                // Quando o lote está cheio, executa o callback
                if (count($currentChunk) >= $chunkSize) {
                    $callback($currentChunk);
                    $currentChunk = []; // Limpa o lote
                }
            }
        }

        // Processa o último lote se tiver sobrado dados
        if (!empty($currentChunk)) {
            $callback($currentChunk);
        }

        $reader->close();
    }
}
