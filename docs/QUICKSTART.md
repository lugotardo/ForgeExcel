# 🚀 Quick Start - ForgeExcel

> **Comece a usar ForgeExcel em 5 minutos!**

---

## Instalação

```bash
composer require lugotardo/forgeexel
```

---

## Exemplo Mais Simples Possível

### 1. Criar um arquivo Excel

```php
<?php
require_once 'vendor/autoload.php';

use Lugotardo\Forgeexel\ForgeExcel;

$dados = [
    ['Nome', 'Email', 'Idade'],
    ['João', 'joao@email.com', 25],
    ['Maria', 'maria@email.com', 30]
];

ForgeExcel::write('pessoas.xlsx', $dados);

echo "Arquivo criado com sucesso!";
```

### 2. Ler um arquivo Excel

```php
<?php
require_once 'vendor/autoload.php';

use Lugotardo\Forgeexel\ForgeExcel;

// Ler como array simples
$dados = ForgeExcel::read('pessoas.xlsx');

foreach ($dados as $linha) {
    echo $linha[0] . ' - ' . $linha[1] . "\n";
}
```

---

## Casos de Uso Comuns

### Exportar do Banco de Dados

```php
// 1. Buscar dados do banco
$usuarios = $pdo->query("SELECT nome, email, idade FROM usuarios")
                ->fetchAll(PDO::FETCH_ASSOC);

// 2. Converter para Excel
$excel = ForgeExcel::arrayToExcel($usuarios);

// 3. Salvar arquivo
ForgeExcel::write('usuarios.xlsx', $excel);
```

### Importar para o Banco de Dados

```php
// 1. Ler Excel com headers
$dados = ForgeExcel::read('clientes.xlsx', true);

// 2. Preparar inserção
$stmt = $pdo->prepare("INSERT INTO clientes (nome, email) VALUES (?, ?)");

// 3. Importar cada linha
foreach ($dados as $cliente) {
    $stmt->execute([$cliente['Nome'], $cliente['Email']]);
}
```

### Criar Relatório Simples

```php
$vendas = [
    ['Produto', 'Quantidade', 'Preço', 'Total'],
    ['Notebook', 2, 3500, '=B2*C2'],
    ['Mouse', 5, 50, '=B3*C3'],
    ['TOTAL', '', '', '=SUM(D2:D3)']
];

ForgeExcel::writeWithFormulas('vendas.xlsx', $vendas);
```

### Criar Tabela Bonita

```php
$funcionarios = [
    ['Nome', 'Cargo', 'Salário'],
    ['João Silva', 'Desenvolvedor', 8500],
    ['Maria Santos', 'Gerente', 12000]
];

// Escolha um tema: blue, green, red, orange, purple
ForgeExcel::writeTable('funcionarios.xlsx', $funcionarios, 'blue');
```

---

## Recursos Principais

### 📖 Leitura

```php
// Leitura simples
$dados = ForgeExcel::read('arquivo.xlsx');

// Com headers (primeira linha vira chave)
$dados = ForgeExcel::read('arquivo.xlsx', true);
// Resultado: [['Nome' => 'João', 'Email' => 'joao@email.com'], ...]

// Apenas primeira aba
$dados = ForgeExcel::readFirstSheet('arquivo.xlsx', true);

// Todas as abas separadas
$abas = ForgeExcel::readAllSheets('arquivo.xlsx', true);

// Contar linhas
$total = ForgeExcel::countRows('arquivo.xlsx');
```

### ✏️ Escrita

```php
// Excel simples
ForgeExcel::write('saida.xlsx', $dados);

// CSV
ForgeExcel::write('saida.csv', $dados, 'csv');

// Com múltiplas abas
$abas = [
    'Aba1' => $dados1,
    'Aba2' => $dados2
];
ForgeExcel::writeWithSheets('arquivo.xlsx', $abas);
```

### 🎨 Formatação

```php
// Estilo simples
$estilos = [
    0 => ['bold' => true, 'background' => '4472C4', 'color' => 'FFFFFF']
];
ForgeExcel::writeWithStyle('arquivo.xlsx', $dados, $estilos);

// Tabela estilizada
ForgeExcel::writeTable('arquivo.xlsx', $dados, 'blue');
```

### 📐 Fórmulas

```php
$dados = [
    ['A', 'B', 'Soma'],
    [10, 20, '=A2+B2'],
    [30, 40, '=A3+B3'],
    ['TOTAL', '', '=SUM(C2:C3)']
];

ForgeExcel::writeWithFormulas('calculos.xlsx', $dados);
```

---

## Dicas Rápidas

### ✅ Para Arquivos Grandes

```php
// Use chunks para não estourar memória
ForgeExcel::readInChunks('arquivo_grande.xlsx', 1000, function($lote) {
    foreach ($lote as $linha) {
        // Processa cada linha
    }
});
```

### ✅ Paleta de Cores

```php
$cores = ForgeExcel::colors();
// Usa: $cores['red'], $cores['blue'], $cores['green'], etc.
```

### ✅ Converter Array Associativo

```php
// De: [['nome' => 'João'], ['nome' => 'Maria']]
// Para: [['nome'], ['João'], ['Maria']]

$excel = ForgeExcel::arrayToExcel($arrayAssociativo);
```

### ✅ Headers Automáticos

```php
// Sem headers
$dados = [
    ['João', 25],
    ['Maria', 30]
];

// Com headers (adicione manualmente)
$dados = [
    ['Nome', 'Idade'],  // Header
    ['João', 25],
    ['Maria', 30]
];
```

---

## Exemplos Práticos de 1 Linha

### Exportar Consulta SQL

```php
ForgeExcel::write('resultado.xlsx', ForgeExcel::arrayToExcel($pdo->query("SELECT * FROM users")->fetchAll(PDO::FETCH_ASSOC)));
```

### Criar Backup de Tabela

```php
ForgeExcel::write("backup_" . date('Y-m-d') . ".xlsx", ForgeExcel::arrayToExcel($pdo->query("SELECT * FROM produtos")->fetchAll(PDO::FETCH_ASSOC)));
```

### Relatório Rápido

```php
ForgeExcel::writeTable('relatorio.xlsx', ForgeExcel::arrayToExcel($dados), 'blue');
```

---

## Próximos Passos

Agora que você já sabe o básico, explore:

1. **[Guia Completo](GUIA_COMPLETO.md)** - Todos os recursos detalhados
2. **[Guia de Formatação](FORMATACAO.md)** - Estilos, cores e formatação avançada
3. **[Guia de Fórmulas](FORMULAS.md)** - Todas as fórmulas suportadas
4. Execute `php test.php` - Ver exemplos funcionando
5. Execute `php test_advanced.php` - Ver recursos avançados

---

## Exemplos Completos

### Sistema de Notas

```php
<?php
require_once 'vendor/autoload.php';
use Lugotardo\Forgeexel\ForgeExcel;

// Ler notas do arquivo
$notas = ForgeExcel::read('notas_entrada.xlsx', true);

// Processar e adicionar média
$resultado = [['Aluno', 'P1', 'P2', 'P3', 'Média', 'Situação']];

foreach ($notas as $aluno) {
    $resultado[] = [
        $aluno['Aluno'],
        $aluno['P1'],
        $aluno['P2'],
        $aluno['P3'],
        '=AVERAGE(B' . (count($resultado) + 1) . ':D' . (count($resultado) + 1) . ')',
        '=IF(E' . (count($resultado) + 1) . '>=7,"Aprovado","Reprovado")'
    ];
}

ForgeExcel::writeWithFormulas('notas_processadas.xlsx', $resultado);
echo "Notas processadas com sucesso!";
```

### Dashboard Executivo

```php
<?php
require_once 'vendor/autoload.php';
use Lugotardo\Forgeexel\ForgeExcel;

// Buscar dados
$pdo = new PDO('mysql:host=localhost;dbname=empresa', 'user', 'pass');

$vendas = $pdo->query("SELECT * FROM vendas_mes")->fetchAll(PDO::FETCH_ASSOC);
$clientes = $pdo->query("SELECT * FROM novos_clientes")->fetchAll(PDO::FETCH_ASSOC);
$produtos = $pdo->query("SELECT * FROM top_produtos LIMIT 10")->fetchAll(PDO::FETCH_ASSOC);

// Criar dashboard com abas
$dashboard = [
    'Resumo' => [
        'data' => [
            ['Métrica', 'Valor'],
            ['Total Vendas', 'R$ 375.000,00'],
            ['Novos Clientes', 127],
            ['Ticket Médio', 'R$ 2.952,00']
        ],
        'headerStyle' => ['bold' => true, 'color' => 'FFFFFF', 'background' => '4472C4']
    ],
    'Vendas' => [
        'data' => ForgeExcel::arrayToExcel($vendas),
        'headerStyle' => ['bold' => true, 'color' => 'FFFFFF', 'background' => '70AD47']
    ],
    'Clientes' => [
        'data' => ForgeExcel::arrayToExcel($clientes),
        'headerStyle' => ['bold' => true, 'color' => 'FFFFFF', 'background' => 'ED7D31']
    ],
    'Top Produtos' => [
        'data' => ForgeExcel::arrayToExcel($produtos),
        'headerStyle' => ['bold' => true, 'color' => 'FFFFFF', 'background' => '7030A0']
    ]
];

ForgeExcel::writeStyledSheets('dashboard_executivo.xlsx', $dashboard);
echo "Dashboard gerado com sucesso!";
```

### Importador em Massa

```php
<?php
require_once 'vendor/autoload.php';
use Lugotardo\Forgeexel\ForgeExcel;

$pdo = new PDO('mysql:host=localhost;dbname=empresa', 'user', 'pass');
$stmt = $pdo->prepare("INSERT INTO produtos (nome, preco, estoque) VALUES (?, ?, ?)");

$total = 0;
$erros = 0;

// Processa 500 linhas por vez
ForgeExcel::readInChunks('produtos_import.xlsx', 500, function($lote) use ($stmt, &$total, &$erros) {
    foreach ($lote as $produto) {
        try {
            $stmt->execute([
                $produto['Nome'],
                $produto['Preco'],
                $produto['Estoque']
            ]);
            $total++;
        } catch (Exception $e) {
            $erros++;
            error_log("Erro: " . $e->getMessage());
        }
    }
    echo "Importados: {$total}, Erros: {$erros}\r";
}, true);

echo "\nImportação concluída!\n";
echo "Total importado: {$total}\n";
echo "Total de erros: {$erros}\n";
```

---

## Troubleshooting Rápido

### Erro: "Arquivo não encontrado"
```php
// Use caminho absoluto
$arquivo = __DIR__ . '/planilhas/dados.xlsx';
ForgeExcel::read($arquivo);
```

### Erro: "Memória esgotada"
```php
// Use chunks
ForgeExcel::readInChunks('arquivo.xlsx', 1000, function($lote) {
    // Processa em lotes
});
```

### Fórmulas não calculam
```php
// Use writeWithFormulas() em vez de write()
ForgeExcel::writeWithFormulas('arquivo.xlsx', $dados);
```

### Caracteres estranhos (����)
```php
// Certifique-se que seus dados estão em UTF-8
$dados = array_map(function($linha) {
    return array_map('utf8_encode', $linha);
}, $dados);
```

---

## Cheat Sheet

| Tarefa | Código |
|--------|--------|
| Criar Excel simples | `ForgeExcel::write('file.xlsx', $data)` |
| Criar CSV | `ForgeExcel::write('file.csv', $data, 'csv')` |
| Ler Excel | `ForgeExcel::read('file.xlsx')` |
| Ler com headers | `ForgeExcel::read('file.xlsx', true)` |
| Converter array | `ForgeExcel::arrayToExcel($array)` |
| Múltiplas abas | `ForgeExcel::writeWithSheets('file.xlsx', $sheets)` |
| Com fórmulas | `ForgeExcel::writeWithFormulas('file.xlsx', $data)` |
| Tabela estilizada | `ForgeExcel::writeTable('file.xlsx', $data, 'blue')` |
| Contar linhas | `ForgeExcel::countRows('file.xlsx')` |
| Processar em lotes | `ForgeExcel::readInChunks('file.xlsx', 1000, $callback)` |

---

## Ajuda

- **Documentação completa:** Veja os arquivos em `docs/`
- **Exemplos funcionando:** Execute `php test.php`
- **Recursos avançados:** Execute `php test_advanced.php`
- **Issues:** Reporte problemas no GitHub

---

**Criado com ❤️ por Luan Gotardo**

Comece simples, evolua gradualmente! 🚀