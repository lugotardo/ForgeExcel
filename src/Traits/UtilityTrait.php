<?php

namespace Lugotardo\Forgeexel\Traits;

/**
 * Trait UtilityTrait
 *
 * Contém todos os métodos utilitários e auxiliares
 *
 * @package Lugotardo\Forgeexel\Traits
 */
trait UtilityTrait
{
    /**
     * Converte um array associativo em array de dados para Excel
     * Útil quando você tem dados de banco de dados ou API
     *
     * Exemplo de uso:
     * $usuarios = [
     *     ['nome' => 'João', 'email' => 'joao@email.com', 'idade' => 25],
     *     ['nome' => 'Maria', 'email' => 'maria@email.com', 'idade' => 30]
     * ];
     * $dados = ForgeExcel::arrayToExcel($usuarios);
     * ForgeExcel::write('usuarios.xlsx', $dados);
     *
     * @param array $associativeArray Array de arrays associativos
     * @param bool $includeHeader Se TRUE, adiciona linha de cabeçalho
     * @return array Array formatado para escrita no Excel
     */
    public static function arrayToExcel(
        array $associativeArray,
        bool $includeHeader = true,
    ): array {
        if (empty($associativeArray)) {
            return [];
        }

        $data = [];

        // Adiciona o cabeçalho (chaves do primeiro item)
        if ($includeHeader) {
            $data[] = array_keys($associativeArray[0]);
        }

        // Adiciona os valores de cada linha
        foreach ($associativeArray as $item) {
            $data[] = array_values($item);
        }

        return $data;
    }
}
