# 📚 Guia Completo - ForgeExcel

> **Documentação completa para manipulação de arquivos Excel em PHP**

---

## 📑 Índice

1. [Introdução](#introdução)
2. [Instalação](#instalação)
3. [Conceitos Básicos](#conceitos-básicos)
4. [Operações Básicas](#operações-básicas)
   - [Leitura de Arquivos](#leitura-de-arquivos)
   - [Escrita de Arquivos](#escrita-de-arquivos)
5. [Recursos Avançados](#recursos-avançados)
   - [Formatação e Estilos](#formatação-e-estilos)
   - [Fórmulas do Excel](#fórmulas-do-excel)
   - [Tabelas Estilizadas](#tabelas-estilizadas)
   - [Múltiplas Abas](#múltiplas-abas)
6. [Referência da API](#referência-da-api)
7. [Exemplos Práticos](#exemplos-práticos)
8. [Melhores Práticas](#melhores-práticas)
9. [Troubleshooting](#troubleshooting)

---

## Introdução

ForgeExcel é uma biblioteca PHP que simplifica drasticamente a manipulação de arquivos Excel. Construída sobre o Box/Spout, oferece uma API intuitiva e poderosa para:

- ✅ Ler e escrever arquivos XLSX, CSV e ODS
- ✅ Aplicar formatação profissional (cores, fontes, negrito)
- ✅ Criar fórmulas do Excel automaticamente
- ✅ Trabalhar com múltiplas abas
- ✅ Processar arquivos gigantes sem estourar memória

---

## Instalação

### Via Composer

```bash
composer require lugotardo/forgeexel
```

### Requisitos

- PHP 7.4 ou superior
- Extensão ZIP habilitada
- Extensão XML habilitada

### Verificação da Instalação

```php
<?php
require_once 'vendor/autoload.php';

use Lugotardo\Forgeexel\ForgeExcel;

echo "ForgeExcel instalado com sucesso!";
```

---

## Conceitos Básicos

### Estrutura de Dados

O ForgeExcel trabalha com arrays PHP simples:

```php
// Array numérico (sem headers)
$dados = [
    ['João', 'joao@email.com', 25],
    ['Maria', 'maria@email.com', 30]
];

// Array associativo (com headers)
$dados = [
    ['nome' => 'João', 'email' => 'joao@email.com', 'idade' => 25],
    ['nome' => 'Maria', 'email' => 'maria@email.com', 'idade' => 30]
];
```

### Formatos Suportados

| Formato | Extensão | Leitura | Escrita | Múltiplas Abas |
|---------|----------|---------|---------|----------------|
| Excel 2007+ | .xlsx | ✅ | ✅ | ✅ |
| CSV | .csv | ✅ | ✅ | ❌ |
| OpenDocument | .ods | ✅ | ✅ | ✅ |

---

## Operações Básicas

### Leitura de Arquivos

#### Leitura Simples

```php
// Retorna array numérico
$dados = ForgeExcel::read('arquivo.xlsx');

foreach ($dados as $linha) {
    echo $linha[0] . ' - ' . $linha[1] . "\n";
}
```

#### Leitura com Headers

```php
// Retorna array associativo usando primeira linha como chave
$dados = ForgeExcel::read('arquivo.xlsx', true);

foreach ($dados as $registro) {
    echo $registro['Nome'] . ' - ' . $registro['Email'] . "\n";
}
```

#### Ler Apenas Primeira Aba

```php
// Mais rápido quando você só precisa da primeira aba
$dados = ForgeExcel::readFirstSheet('arquivo.xlsx', true);
```

#### Ler Todas as Abas Separadamente

```php
$todasAbas = ForgeExcel::readAllSheets('arquivo.xlsx', true);

foreach ($todasAbas as $nomeAba => $dados) {
    echo "Aba: {$nomeAba}\n";
    echo "Total de registros: " . count($dados) . "\n";
}
```

#### Contar Linhas

```php
$total = ForgeExcel::countRows('arquivo.xlsx');
$totalSemHeader = ForgeExcel::countRows('arquivo.xlsx', false);

echo "Total: {$total} linhas\n";
```

#### Leitura em Lotes (Arquivos Grandes)

```php
// Processa 1000 linhas por vez (ideal para arquivos gigantes)
ForgeExcel::readInChunks('arquivo_grande.xlsx', 1000, function($lote) {
    foreach ($lote as $linha) {
        // Processa cada linha
        // Salva no banco, envia email, etc
        processarLinha($linha);
    }
}, true);
```

---

### Escrita de Arquivos

#### Escrita Simples

```php
$dados = [
    ['Nome', 'Email', 'Idade'],
    ['João Silva', 'joao@email.com', 25],
    ['Maria Santos', 'maria@email.com', 30]
];

ForgeExcel::write('saida.xlsx', $dados);
```

#### Criar CSV

```php
$dados = [
    ['Produto', 'Preço'],
    ['Notebook', 3500.00],
    ['Mouse', 50.00]
];

ForgeExcel::write('produtos.csv', $dados, 'csv');
```

#### Converter Array Associativo

```php
// Dados do banco de dados
$usuarios = $pdo->query("SELECT nome, email, idade FROM usuarios")
                ->fetchAll(PDO::FETCH_ASSOC);

// Converte e adiciona headers automaticamente
$dadosExcel = ForgeExcel::arrayToExcel($usuarios);

ForgeExcel::write('usuarios.xlsx', $dadosExcel);
```

#### Múltiplas Abas

```php
$abas = [
    'Clientes' => [
        ['ID', 'Nome', 'Email'],
        [1, 'João', 'joao@email.com'],
        [2, 'Maria', 'maria@email.com']
    ],
    'Produtos' => [
        ['Código', 'Produto', 'Preço'],
        ['A001', 'Notebook', 3500.00],
        ['A002', 'Mouse', 50.00]
    ]
];

ForgeExcel::writeWithSheets('relatorio.xlsx', $abas);
```

---

## Recursos Avançados

### Formatação e Estilos

#### Criar Estilo Personalizado

```php
$estilo = ForgeExcel::createStyle([
    'bold' => true,              // Negrito
    'italic' => true,            // Itálico
    'underline' => true,         // Sublinhado
    'fontSize' => 14,            // Tamanho da fonte
    'fontName' => 'Arial',       // Nome da fonte
    'color' => 'FF0000',         // Cor do texto (hex sem #)
    'background' => 'FFFF00',    // Cor de fundo (hex sem #)
    'align' => 'center',         // Alinhamento (left, center, right)
    'wrapText' => true,          // Quebrar texto automaticamente
    'border' => true             // Adicionar bordas
]);
```

#### Escrever com Formatação

```php
$dados = [
    ['Produto', 'Quantidade', 'Preço', 'Total'],
    ['Notebook', 5, 3500.00, 17500.00],
    ['Mouse', 25, 80.00, 2000.00]
];

// Estilos por linha
$estilosPorLinha = [
    0 => ['bold' => true, 'color' => 'FFFFFF', 'background' => '4472C4'], // Header
];

// Estilos por coluna
$estilosPorColuna = [
    2 => ['align' => 'right'],  // Coluna de preço
    3 => ['align' => 'right', 'bold' => true]  // Coluna de total
];

ForgeExcel::writeWithStyle('vendas.xlsx', $dados, $estilosPorLinha, $estilosPorColuna);
```

#### Paleta de Cores

```php
// Obter todas as cores disponíveis
$cores = ForgeExcel::colors();

// Usar cores predefinidas
$estilo = ForgeExcel::createStyle([
    'color' => $cores['red'],
    'background' => $cores['light_gray']
]);

// Cores disponíveis:
// black, white, red, green, blue, yellow, orange, purple, pink,
// gray, light_gray, dark_gray, cyan, magenta, lime, navy, teal,
// olive, maroon, aqua
```

#### Alinhamentos

```php
$alinhamentos = ForgeExcel::alignments();

$estilo = ForgeExcel::createStyle([
    'align' => $alinhamentos['center']  // left, center, right
]);
```

---

### Fórmulas do Excel

#### Fórmulas Básicas

```php
$dados = [
    ['Produto', 'Quantidade', 'Preço', 'Total'],
    ['Notebook', 2, 3500, '=B2*C2'],     // Multiplicação
    ['Mouse', 5, 50, '=B3*C3'],
    ['TOTAL', '', '', '=SUM(D2:D3)']     // Soma
];

ForgeExcel::writeWithFormulas('vendas.xlsx', $dados);
```

#### Fórmulas Avançadas

```php
$dados = [
    ['Mês', 'Receita', 'Despesas', 'Lucro', 'Margem %'],
    ['Janeiro', 50000, 35000, '=B2-C2', '=(D2/B2)*100'],
    ['Fevereiro', 62000, 42000, '=B3-C3', '=(D3/B3)*100'],
    ['', '', '', '', ''],
    ['TOTAL', '=SUM(B2:B3)', '=SUM(C2:C3)', '=SUM(D2:D3)', '=AVERAGE(E2:E3)'],
    ['MÉDIA', '=AVERAGE(B2:B3)', '=AVERAGE(C2:C3)', '=AVERAGE(D2:D3)', '']
];

$headerStyle = ['bold' => true, 'background' => '4472C4', 'color' => 'FFFFFF'];

ForgeExcel::writeWithFormulas('financeiro.xlsx', $dados, $headerStyle);
```

#### Fórmulas Condicionais

```php
$dados = [
    ['Aluno', 'Nota', 'Situação'],
    ['João', 8.5, '=IF(B2>=7,"Aprovado","Reprovado")'],
    ['Maria', 6.0, '=IF(B3>=7,"Aprovado","Reprovado")'],
    ['Pedro', 7.5, '=IF(B4>=7,"Aprovado","Reprovado")']
];

ForgeExcel::writeWithFormulas('notas.xlsx', $dados);
```

#### Fórmulas com Referências Absolutas

```php
$dados = [
    ['Item', 'Valor', '% do Total'],
    ['Item 1', 100, '=B2/$B$5*100'],
    ['Item 2', 200, '=B3/$B$5*100'],
    ['Item 3', 150, '=B4/$B$5*100'],
    ['TOTAL', '=SUM(B2:B4)', '100%']
];

ForgeExcel::writeWithFormulas('percentuais.xlsx', $dados);
```

#### Fórmulas Suportadas

| Fórmula | Descrição | Exemplo |
|---------|-----------|---------|
| SUM | Soma valores | `=SUM(A1:A10)` |
| AVERAGE | Média | `=AVERAGE(A1:A10)` |
| COUNT | Conta células | `=COUNT(A1:A10)` |
| COUNTIF | Conta com condição | `=COUNTIF(A1:A10,"Ativo")` |
| IF | Condicional | `=IF(A1>10,"Alto","Baixo")` |
| MAX | Valor máximo | `=MAX(A1:A10)` |
| MIN | Valor mínimo | `=MIN(A1:A10)` |
| ROUND | Arredondamento | `=ROUND(A1,2)` |
| CONCATENATE | Concatenar textos | `=CONCATENATE(A1," ",B1)` |

---

### Tabelas Estilizadas

#### Temas Disponíveis

```php
// 5 temas predefinidos: blue, green, red, orange, purple

$dados = [
    ['Nome', 'Cargo', 'Salário'],
    ['João Silva', 'Desenvolvedor', 8500],
    ['Maria Santos', 'Gerente', 12000]
];

ForgeExcel::writeTable('funcionarios.xlsx', $dados, 'blue');
```

#### Exemplos de Temas

**Tema Blue** (Azul Profissional)
```php
ForgeExcel::writeTable('arquivo.xlsx', $dados, 'blue');
```
- Header: Azul escuro com texto branco
- Linhas alternadas: Azul claro / Branco

**Tema Green** (Verde Natureza)
```php
ForgeExcel::writeTable('arquivo.xlsx', $dados, 'green');
```
- Header: Verde com texto branco
- Linhas alternadas: Verde claro / Branco

**Tema Red** (Vermelho Corporativo)
```php
ForgeExcel::writeTable('arquivo.xlsx', $dados, 'red');
```
- Header: Vermelho alaranjado com texto branco
- Linhas alternadas: Laranja claro / Branco

**Tema Orange** (Laranja Vibrante)
```php
ForgeExcel::writeTable('arquivo.xlsx', $dados, 'orange');
```
- Header: Laranja com texto branco
- Linhas alternadas: Bege / Branco

**Tema Purple** (Roxo Elegante)
```php
ForgeExcel::writeTable('arquivo.xlsx', $dados, 'purple');
```
- Header: Roxo escuro com texto branco
- Linhas alternadas: Lilás claro / Branco

---

### Múltiplas Abas com Estilos

```php
$abas = [
    'Dashboard' => [
        'data' => [
            ['Métrica', 'Valor'],
            ['Total de Vendas', 375000],
            ['Novos Clientes', 127]
        ],
        'headerStyle' => [
            'bold' => true,
            'color' => 'FFFFFF',
            'background' => '4472C4',
            'fontSize' => 13
        ]
    ],
    'Detalhes' => [
        'data' => [
            ['Produto', 'Quantidade', 'Total'],
            ['Notebook', 45, '=B2*3500'],
            ['Mouse', 120, '=B3*80']
        ],
        'headerStyle' => [
            'bold' => true,
            'color' => 'FFFFFF',
            'background' => '70AD47'
        ],
        'rowStyles' => [
            1 => ['background' => 'D9E1F2'],
            2 => ['background' => 'E2EFDA']
        ]
    ]
];

ForgeExcel::writeStyledSheets('relatorio_completo.xlsx', $abas);
```

---

## Referência da API

### Métodos de Leitura

#### `read(string $filePath, bool $firstRowAsHeader = false): array`

Lê um arquivo Excel completo.

**Parâmetros:**
- `$filePath`: Caminho do arquivo
- `$firstRowAsHeader`: Se TRUE, primeira linha vira chave do array

**Retorna:** Array com todos os dados

**Exemplo:**
```php
$dados = ForgeExcel::read('arquivo.xlsx', true);
```

---

#### `readFirstSheet(string $filePath, bool $firstRowAsHeader = false): array`

Lê apenas a primeira aba (mais rápido).

**Parâmetros:**
- `$filePath`: Caminho do arquivo
- `$firstRowAsHeader`: Usar primeira linha como header

**Retorna:** Array com dados da primeira aba

**Exemplo:**
```php
$dados = ForgeExcel::readFirstSheet('arquivo.xlsx', true);
```

---

#### `readAllSheets(string $filePath, bool $firstRowAsHeader = false): array`

Lê todas as abas separadamente.

**Parâmetros:**
- `$filePath`: Caminho do arquivo
- `$firstRowAsHeader`: Usar primeira linha como header

**Retorna:** Array associativo [nome_aba => dados]

**Exemplo:**
```php
$todasAbas = ForgeExcel::readAllSheets('arquivo.xlsx', true);
foreach ($todasAbas as $nomeAba => $dados) {
    // Processa cada aba
}
```

---

#### `readInChunks(string $filePath, int $chunkSize, callable $callback, bool $firstRowAsHeader = false): void`

Processa arquivo em lotes (ideal para arquivos grandes).

**Parâmetros:**
- `$filePath`: Caminho do arquivo
- `$chunkSize`: Tamanho do lote (ex: 1000)
- `$callback`: Função que recebe cada lote
- `$firstRowAsHeader`: Usar primeira linha como header

**Exemplo:**
```php
ForgeExcel::readInChunks('arquivo.xlsx', 1000, function($lote) {
    foreach ($lote as $linha) {
        processarLinha($linha);
    }
}, true);
```

---

#### `countRows(string $filePath, bool $countHeader = true): int`

Conta linhas do arquivo.

**Parâmetros:**
- `$filePath`: Caminho do arquivo
- `$countHeader`: Incluir header na contagem

**Retorna:** Número de linhas

**Exemplo:**
```php
$total = ForgeExcel::countRows('arquivo.xlsx');
```

---

### Métodos de Escrita

#### `write(string $filePath, array $data, string $type = 'xlsx'): bool`

Escreve dados em arquivo.

**Parâmetros:**
- `$filePath`: Caminho de saída
- `$data`: Array de dados
- `$type`: Tipo do arquivo (xlsx, csv, ods)

**Retorna:** TRUE se sucesso

**Exemplo:**
```php
ForgeExcel::write('saida.xlsx', $dados);
ForgeExcel::write('saida.csv', $dados, 'csv');
```

---

#### `writeWithSheets(string $filePath, array $sheets): bool`

Cria arquivo com múltiplas abas.

**Parâmetros:**
- `$filePath`: Caminho de saída
- `$sheets`: Array [nome_aba => dados]

**Retorna:** TRUE se sucesso

**Exemplo:**
```php
$abas = [
    'Aba1' => $dados1,
    'Aba2' => $dados2
];
ForgeExcel::writeWithSheets('arquivo.xlsx', $abas);
```

---

#### `writeWithStyle(string $filePath, array $data, array $styles = [], array $columnStyles = []): bool`

Escreve com formatação personalizada.

**Parâmetros:**
- `$filePath`: Caminho de saída
- `$data`: Array de dados
- `$styles`: Estilos por linha [numero_linha => opcoes]
- `$columnStyles`: Estilos por coluna [numero_coluna => opcoes]

**Retorna:** TRUE se sucesso

**Exemplo:**
```php
$estilos = [
    0 => ['bold' => true, 'background' => '4472C4']
];
ForgeExcel::writeWithStyle('arquivo.xlsx', $dados, $estilos);
```

---

#### `writeWithFormulas(string $filePath, array $data, array $headerStyle = []): bool`

Escreve dados com fórmulas.

**Parâmetros:**
- `$filePath`: Caminho de saída
- `$data`: Array com fórmulas (strings iniciando com =)
- `$headerStyle`: Estilo opcional do header

**Retorna:** TRUE se sucesso

**Exemplo:**
```php
$dados = [
    ['A', 'B', 'Total'],
    [10, 20, '=A2+B2']
];
ForgeExcel::writeWithFormulas('arquivo.xlsx', $dados);
```

---

#### `writeTable(string $filePath, array $data, string $theme = 'blue'): bool`

Cria tabela com tema predefinido.

**Parâmetros:**
- `$filePath`: Caminho de saída
- `$data`: Array de dados
- `$theme`: Tema (blue, green, red, orange, purple)

**Retorna:** TRUE se sucesso

**Exemplo:**
```php
ForgeExcel::writeTable('tabela.xlsx', $dados, 'blue');
```

---

#### `writeStyledSheets(string $filePath, array $sheets): bool`

Cria múltiplas abas com estilos.

**Parâmetros:**
- `$filePath`: Caminho de saída
- `$sheets`: Array [nome_aba => ['data' => dados, 'headerStyle' => estilo]]

**Retorna:** TRUE se sucesso

**Exemplo:**
```php
$abas = [
    'Aba1' => [
        'data' => $dados,
        'headerStyle' => ['bold' => true]
    ]
];
ForgeExcel::writeStyledSheets('arquivo.xlsx', $abas);
```

---

### Métodos Utilitários

#### `arrayToExcel(array $associativeArray, bool $includeHeader = true): array`

Converte array associativo para formato Excel.

**Parâmetros:**
- `$associativeArray`: Array de arrays associativos
- `$includeHeader`: Incluir linha de cabeçalho

**Retorna:** Array formatado

**Exemplo:**
```php
$usuarios = [
    ['nome' => 'João', 'email' => 'joao@email.com']
];
$excel = ForgeExcel::arrayToExcel($usuarios);
```

---

#### `createStyle(array $options = []): Style`

Cria objeto de estilo personalizado.

**Parâmetros:**
- `$options`: Array de opções de estilo

**Retorna:** Objeto Style do Spout

**Opções disponíveis:**
- `bold`: Negrito (boolean)
- `italic`: Itálico (boolean)
- `underline`: Sublinhado (boolean)
- `fontSize`: Tamanho da fonte (int)
- `fontName`: Nome da fonte (string)
- `color`: Cor do texto hex sem # (string)
- `background`: Cor de fundo hex sem # (string)
- `align`: Alinhamento (left, center, right)
- `wrapText`: Quebrar texto (boolean)
- `border`: Adicionar bordas (boolean)

**Exemplo:**
```php
$estilo = ForgeExcel::createStyle([
    'bold' => true,
    'color' => 'FFFFFF',
    'background' => '4472C4'
]);
```

---

#### `colors(): array`

Retorna array com cores predefinidas.

**Retorna:** Array [nome => hex]

**Exemplo:**
```php
$cores = ForgeExcel::colors();
echo $cores['red']; // FF0000
```

---

#### `alignments(): array`

Retorna constantes de alinhamento.

**Retorna:** Array [nome => constante]

**Exemplo:**
```php
$align = ForgeExcel::alignments();
$estilo = ForgeExcel::createStyle(['align' => $align['center']]);
```

---

## Exemplos Práticos

### 1. Importar Clientes de Planilha

```php
<?php
// Ler planilha de clientes
$clientes = ForgeExcel::read('clientes.xlsx', true);

// Importar para o banco de dados
$stmt = $pdo->prepare("INSERT INTO clientes (nome, email, telefone) VALUES (?, ?, ?)");

foreach ($clientes as $cliente) {
    $stmt->execute([
        $cliente['Nome'],
        $cliente['Email'],
        $cliente['Telefone']
    ]);
}

echo "Importados " . count($clientes) . " clientes!";
```

---

### 2. Exportar Relatório de Vendas

```php
<?php
// Buscar vendas do banco
$vendas = $pdo->query("
    SELECT 
        v.data,
        c.nome as cliente,
        p.nome as produto,
        v.quantidade,
        v.valor_total
    FROM vendas v
    JOIN clientes c ON v.cliente_id = c.id
    JOIN produtos p ON v.produto_id = p.id
    WHERE MONTH(v.data) = MONTH(CURRENT_DATE)
")->fetchAll(PDO::FETCH_ASSOC);

// Converter e exportar
$dados = ForgeExcel::arrayToExcel($vendas);
ForgeExcel::write('relatorio_vendas.xlsx', $dados);

echo "Relatório gerado com sucesso!";
```

---

### 3. Dashboard Executivo

```php
<?php
// Calcular métricas
$totalVendas = $pdo->query("SELECT SUM(valor) FROM vendas")->fetchColumn();
$totalClientes = $pdo->query("SELECT COUNT(*) FROM clientes")->fetchColumn();
$ticketMedio = $totalVendas / $totalClientes;

// Top produtos
$topProdutos = $pdo->query("
    SELECT nome, SUM(quantidade) as qtd, SUM(valor_total) as total
    FROM vendas v
    JOIN produtos p ON v.produto_id = p.id
    GROUP BY p.id
    ORDER BY total DESC
    LIMIT 10
")->fetchAll(PDO::FETCH_ASSOC);

// Vendas por mês
$vendasMes = $pdo->query("
    SELECT 
        DATE_FORMAT(data, '%Y-%m') as mes,
        COUNT(*) as quantidade,
        SUM(valor_total) as total
    FROM vendas
    GROUP BY DATE_FORMAT(data, '%Y-%m')
    ORDER BY mes DESC
    LIMIT 12
")->fetchAll(PDO::FETCH_ASSOC);

// Criar dashboard
$abas = [
    'Resumo' => [
        'data' => [
            ['Métrica', 'Valor'],
            ['Total de Vendas', 'R$ ' . number_format($totalVendas, 2, ',', '.')],
            ['Total de Clientes', $totalClientes],
            ['Ticket Médio', 'R$ ' . number_format($ticketMedio, 2, ',', '.')]
        ],
        'headerStyle' => ['bold' => true, 'color' => 'FFFFFF', 'background' => '4472C4']
    ],
    'Top 10 Produtos' => [
        'data' => ForgeExcel::arrayToExcel($topProdutos),
        'headerStyle' => ['bold' => true, 'color' => 'FFFFFF', 'background' => '70AD47']
    ],
    'Vendas por Mês' => [
        'data' => ForgeExcel::arrayToExcel($vendasMes),
        'headerStyle' => ['bold' => true, 'color' => 'FFFFFF', 'background' => 'ED7D31']
    ]
];

ForgeExcel::writeStyledSheets('dashboard_executivo.xlsx', $abas);

echo "Dashboard gerado com sucesso!";
```

---

### 4. Processar Arquivo Enorme

```php
<?php
// Processar arquivo de 1 milhão de linhas
$totalProcessado = 0;
$totalErros = 0;

ForgeExcel::readInChunks('arquivo_grande.xlsx', 1000, function($lote) use (&$totalProcessado, &$totalErros, $pdo) {
    
    $stmt = $pdo->prepare("INSERT INTO dados (campo1, campo2, campo3) VALUES (?, ?, ?)");
    
    foreach ($lote as $linha) {
        try {
            $stmt->execute([$linha[0], $linha[1], $linha[2]]);
            $totalProcessado++;
        } catch (Exception $e) {
            $totalErros++;
            error_log("Erro na linha: " . $e->getMessage());
        }
    }
    
    echo "Processadas {$totalProcessado} linhas...\n";
}, true);

echo "\nProcessamento concluído!\n";
echo "Total processado: {$totalProcessado}\n";
echo "Total de erros: {$totalErros}\n";
```

---

### 5. Nota Fiscal com Fórmulas

```php
<?php
$itens = [
    ['Item', 'Quantidade', 'Valor Unit.', 'Subtotal', 'IPI (10%)', 'Total'],
    ['Produto A', 10, 100.00, '=B2*C2', '=D2*0.1', '=D2+E2'],
    ['Produto B', 5, 200.00, '=B3*C3', '=D3*0.1', '=D3+E3'],
    ['Produto C', 15, 50.00, '=B4*C4', '=D4*0.1', '=D4+E4'],
    ['', '', '', '', '', ''],
    ['TOTAIS', '=SUM(B2:B4)', '', '=SUM(D2:D4)', '=SUM(E2:E4)', '=SUM(F2:F4)']
];

$headerStyle = [
    'bold' => true,
    'color' => 'FFFFFF',
    'background' => '203864',
    'fontSize' => 11
];

ForgeExcel::writeWithFormulas('nota_fiscal.xlsx', $itens, $headerStyle);

echo "Nota fiscal gerada com fórmulas automáticas!";
```

---

### 6. Comparar Duas Planilhas

```php
<?php
// Ler duas planilhas
$arquivo1 = ForgeExcel::read('planilha1.xlsx', true);
$arquivo2 = ForgeExcel::read('planilha2.xlsx', true);

// Criar arrays associativos por ID
$dados1 = array_column($arquivo1, null, 'ID');
$dados2 = array_column($arquivo2, null, 'ID');

// Encontrar diferenças
$diferencas = [];
$diferencas[] = ['ID', 'Campo', 'Valor Arquivo 1', 'Valor Arquivo 2'];

foreach ($dados1 as $id => $registro1) {
    if (isset($dados2[$id])) {
        $registro2 = $dados2[$id];
        
        foreach ($registro1 as $campo => $valor1) {
            $valor2 = $registro2[$campo] ?? '';
            
            if ($valor1 != $valor2) {
                $diferencas[] = [$id, $campo, $valor1, $valor2];
            }
        }
    }
}

// Salvar relatório de diferenças
ForgeExcel::writeTable('diferencas.xlsx', $diferencas, 'red');

echo "Encontradas " . (count($diferencas) - 1) . " diferenças!";
```

---

### 7. Gerar Folha de Pagamento

```php
<?php
$funcionarios = [
    ['Nome', 'Salário Base', 'Horas Extra', 'Valor Hora Extra', 'Total Extras', 'Total Bruto', 'INSS (11%)', 'IRRF (15%)', 'Total Líquido'],
    ['João Silva', 3000, 10, 25, '=C2*D2', '=B2+E2', '=F2*0.11', '=F2*0.15', '=F2-G2-H2'],
    ['Maria Santos', 4500, 5, 37.50, '=C3*D3', '=B3+E3', '=F3*0.11', '=F3*0.15', '=F3-G3-H3'],
    ['Pedro Costa', 5000, 8, 41.67, '=C4*D4', '=B4+E4', '=F4*0.11', '=F4*0.15', '=F4-G4-H4'],
    ['', '', '', '', '', '', '', '', ''],
    ['TOTAIS', '=SUM(B2:B4)', '=SUM(C2:C4)', '', '=SUM(E2:E4)', '=SUM(F2:F4)', '=SUM(G2:G4)', '=SUM(H2:H4)', '=SUM(I2:I4)']
];

$headerStyle = [
    'bold' => true,
    'color' => 'FFFFFF',
    'background' => '203864',
    'fontSize' => 10,
    'wrapText' => true
];

ForgeExcel::writeWithFormulas('folha_pagamento.xlsx', $funcionarios, $headerStyle);

echo "Folha de pagamento gerada!";
```

---

## Melhores Práticas

### 1. Performance

#### ✅ Use `readInChunks()` para arquivos grandes

```php
// BOM: Processa em lotes
ForgeExcel::readInChunks('arquivo_grande.xlsx', 1000, function($lote) {
    processarLote($lote);
});

// RUIM: Carrega tudo na memória
$todosOsDados = ForgeExcel::read('arquivo_grande.xlsx');
```

#### ✅ Use `readFirstSheet()` quando possível

```php
// BOM: Mais rápido se só precisa da primeira aba
$dados = ForgeExcel::readFirstSheet('arquivo.xlsx');

// RUIM: Lê todas as abas desnecessariamente
$dados = ForgeExcel::read('arquivo.xlsx');
```

#### ✅ Conte linhas antes de processar

```php
// BOM: Sabe o que esperar
$total = ForgeExcel::countRows('arquivo.xlsx');
echo "Processando {$total} linhas...\n";

$dados = ForgeExcel::read('arquivo.xlsx');
```

---

### 2. Tratamento de Erros

#### ✅ Sempre use try-catch

```php
try {
    $dados = ForgeExcel::read('arquivo.xlsx');
    processar($dados);
} catch (Exception $e) {
    error_log("Erro ao ler arquivo: " . $e->getMessage());
    // Lidar com o erro apropriadamente
}
```

#### ✅ Valide dados antes de processar

```php
$dados = ForgeExcel::read('arquivo.xlsx', true);

foreach ($dados as $linha) {
    // Valida campos obrigatórios
    if (empty($linha['Email']) || !filter_var($linha['Email'], FILTER_VALIDATE_EMAIL)) {
        error_log("Email inválido: " . $linha['Email']);
        continue;
    }
    
    // Processa linha válida
    processarLinha($linha);
}
```

---

### 3. Organização de Código

#### ✅ Use funções auxiliares

```php
function exportarRelatorio($dados, $nomeArquivo) {
    $excel = ForgeExcel::arrayToExcel($dados);
    ForgeExcel::writeTable($nomeArquivo, $excel, 'blue');
    return $nomeArquivo;
}

// Uso
$vendas = buscarVendasDoBanco();
$arquivo = exportarRelatorio($vendas, 'vendas_' . date('Y-m-d') . '.xlsx');
```

#### ✅ Crie templates de estilo

```php
class ExcelStyles {
    public static function header() {
        return [
            'bold' => true,
            'color' => 'FFFFFF',
            'background' => '4472C4',
            'fontSize' => 12
        ];
    }
    
    public static function totals() {
        return [
            'bold' => true,
            'background' => 'FFFF00'
        ];
    }
}

// Uso
ForgeExcel::writeWithStyle($arquivo, $dados, [
    0 => ExcelStyles::header(),
    10 => ExcelStyles::totals()
]);
```

---

### 4. Segurança

#### ✅ Valide tipo de arquivo

```php
$allowedExtensions = ['xlsx', 'csv', 'ods'];
$extension = pathinfo($_FILES['file']['name'], PATHINFO_EXTENSION);

if (!in_array($extension, $allowedExtensions)) {
    throw new Exception("Tipo de arquivo não permitido");
}
```

#### ✅ Limite tamanho de arquivo

```php
$maxSize = 10 * 1024 * 1024; // 10MB

if ($_FILES['file']['size'] > $maxSize) {
    throw new Exception("Arquivo muito grande");
}
```

#### ✅ Sanitize dados de entrada

```php
$dados = ForgeExcel::read($arquivo, true);

foreach ($dados as $linha) {
    $nome = htmlspecialchars($linha['Nome'], ENT_QUOTES, 'UTF-8');
    $email = filter_var($linha['Email'], FILTER_SANITIZE_EMAIL);
    
    // Use dados sanitizados
}
```

---

## Troubleshooting

### Problema: Memória esgotada

**Sintoma:**
```
PHP Fatal error: Allowed memory size exhausted
```

**Solução:**
Use `readInChunks()` para processar em lotes:

```php
ForgeExcel::readInChunks('arquivo.xlsx', 1000, function($lote) {
    // Processa 1000 linhas por vez
});
```

---

### Problema: Arquivo não encontrado

**Sintoma:**
```
Exception: Arquivo não encontrado
```

**Solução:**
Verifique o caminho e permissões:

```php
$arquivo = __DIR__ . '/planilhas/dados.xlsx';

if (!file_exists($arquivo)) {
    die("Arquivo não existe: {$arquivo}");
}

if (!is_readable($arquivo)) {
    die("Arquivo não pode ser lido: {$arquivo}");
}
```

---

### Problema: Caracteres especiais corrompidos

**Sintoma:**
Acentos aparecem como �����

**Solução:**
Use UTF-8 em todo fluxo:

```php
// Ao escrever
$dados = [
    ['Nome', 'Descrição'],
    ['João', 'Programação'],
];

ForgeExcel::write('arquivo.xlsx', $dados);

// Ao ler
$dados = ForgeExcel::read('arquivo.xlsx');
// Dados já vêm em UTF-8
```

---

### Problema: Fórmulas não calculam

**Sintoma:**
Fórmulas aparecem como texto

**Solução:**
Use `writeWithFormulas()` em vez de `write()`:

```php
// CORRETO
ForgeExcel::writeWithFormulas('arquivo.xlsx', $dados);

// INCORRETO
ForgeExcel::write('arquivo.xlsx', $dados);
```

---

### Problema: Estilos não aplicados

**Sintoma:**
Arquivo criado mas sem formatação

**Solução:**
Use os métodos específicos de estilo:

```php
// CORRETO
ForgeExcel::writeWithStyle('arquivo.xlsx', $dados, $estilos);

// INCORRETO
ForgeExcel::write('arquivo.xlsx', $dados);
```

---

### Problema: CSV com separador errado

**Sintoma:**
CSV abre tudo em uma coluna

**Solução:**
O Spout usa vírgula por padrão (padrão internacional). Para Excel brasileiro, abra manualmente escolhendo o delimitador.

---

### Problema: Arquivo muito lento para processar

**Sintoma:**
Processamento demora muito

**Solução:**

1. Use chunks para arquivos grandes
2. Processe em background se possível
3. Mostre progresso ao usuário

```php
$total = ForgeExcel::countRows('arquivo.xlsx', false);
$processados = 0;

ForgeExcel::readInChunks('arquivo.xlsx', 500, function($lote) use ($total, &$processados) {
    foreach ($lote as $linha) {
        processarLinha($linha);
        $processados++;
        
        if ($processados % 100 === 0) {
            $percentual = ($processados / $total) * 100;
            echo "Progresso: " . round($percentual, 2) . "%\n";
        }
    }
}, true);
```

---

## Recursos Adicionais

### Links Úteis

- **Repositório GitHub:** [github.com/lugotardo/forgeexcel](https://github.com)
- **Documentação Box/Spout:** [box/spout](https://github.com/box/spout)
- **Issues e Suporte:** [GitHub Issues](https://github.com)

### Contribuindo

Contribuições são bem-vindas! Por favor:

1. Fork o projeto
2. Crie uma branch para sua feature
3. Commit suas mudanças
4. Push para a branch
5. Abra um Pull Request

### Licença

Este projeto está sob a licença MIT.

---

**Desenvolvido com ❤️ por Luan Gotardo**

Para dúvidas ou sugestões: luan.gotardo.dev@gmail.com