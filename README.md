# 📊 ForgeExcel

**A maneira mais fácil de trabalhar com arquivos Excel em PHP!**

ForgeExcel é uma biblioteca PHP que simplifica a leitura e escrita de arquivos Excel (XLSX, CSV, ODS), mas com uma interface muito mais simples e intuitiva.

## 🚀 Instalação

```bash
composer require lugotardo/forgeexel
```

## ✨ Características

- ✅ **Super Simples**: API limpa e fácil de usar
- 📖 **Leitura Rápida**: Lê arquivos Excel em segundos
- ✏️ **Escrita Fácil**: Cria arquivos Excel com poucas linhas
- 📑 **Múltiplas Abas**: Suporte para arquivos com várias abas
- 🔄 **Arrays Associativos**: Conversão automática de dados
- 💾 **Memória Eficiente**: Processa arquivos grandes em lotes (chunks)
- 📊 **Múltiplos Formatos**: XLSX, CSV, ODS
- 🇧🇷 **Documentação em Português**: Comentários explicativos em português
- 🎨 **Formatação Profissional**: Negrito, cores, fontes, alinhamento e bordas
- 📐 **Fórmulas do Excel**: Crie planilhas com cálculos automáticos
- 🎯 **Tabelas Estilizadas**: 5 temas prontos para usar

## 📚 Uso Básico

### Criar um arquivo Excel simples

```php
use Lugotardo\Forgeexel\ForgeExcel;

$dados = [
    ['Nome', 'Email', 'Idade'],
    ['João Silva', 'joao@email.com', 25],
    ['Maria Santos', 'maria@email.com', 30],
];

ForgeExcel::write('pessoas.xlsx', $dados);
```

### Ler um arquivo Excel

```php
// Leitura simples (array numérico)
$dados = ForgeExcel::read('pessoas.xlsx');

// Leitura com headers (array associativo)
$dados = ForgeExcel::read('pessoas.xlsx', true);

foreach ($dados as $pessoa) {
    echo $pessoa['Nome'] . ' - ' . $pessoa['Email'] . "\n";
}
```

## 🎯 Exemplos Práticos

### Exportar dados do banco para Excel

```php
// Simula dados do banco
$usuarios = $pdo->query("SELECT nome, email, idade FROM usuarios")->fetchAll(PDO::FETCH_ASSOC);

// Converte e salva em Excel (headers automáticos!)
$dadosExcel = ForgeExcel::arrayToExcel($usuarios);
ForgeExcel::write('usuarios.xlsx', $dadosExcel);
```

### Criar Excel com múltiplas abas

```php
$relatorio = [
    'Clientes' => [
        ['ID', 'Nome', 'Email'],
        [1, 'João', 'joao@email.com'],
        [2, 'Maria', 'maria@email.com'],
    ],
    'Produtos' => [
        ['Código', 'Produto', 'Preço'],
        ['A001', 'Notebook', 3500.00],
        ['A002', 'Mouse', 50.00],
    ],
    'Vendas' => [
        ['Data', 'Cliente', 'Valor'],
        ['2024-01-15', 'João', 3500.00],
        ['2024-01-16', 'Maria', 50.00],
    ]
];

ForgeExcel::writeWithSheets('relatorio.xlsx', $relatorio);
```

### Processar arquivos grandes (sem estourar memória)

```php
// Processa arquivo enorme em lotes de 100 linhas
ForgeExcel::readInChunks('arquivo_gigante.xlsx', 100, function($lote) {
    foreach ($lote as $linha) {
        // Processa cada linha
        // Salva no banco, envia email, etc
        processarLinha($linha);
    }
});
```

### Ler apenas a primeira aba

```php
// Mais rápido quando você só precisa da primeira aba
$dados = ForgeExcel::readFirstSheet('arquivo.xlsx', true);
```

### Ler todas as abas separadamente

```php
$todasAbas = ForgeExcel::readAllSheets('arquivo.xlsx', true);

foreach ($todasAbas as $nomeAba => $dados) {
    echo "Aba: {$nomeAba} tem " . count($dados) . " registros\n";
    
    foreach ($dados as $linha) {
        // Processa cada linha de cada aba
    }
}
```

### Criar arquivo CSV

```php
$dados = [
    ['Produto', 'Quantidade', 'Preço'],
    ['Caneta', 100, 2.50],
    ['Caderno', 50, 15.00],
];

ForgeExcel::write('estoque.csv', $dados, 'csv');
```

### Contar linhas de um arquivo

```php
$total = ForgeExcel::countRows('arquivo.xlsx');
$totalSemHeader = ForgeExcel::countRows('arquivo.xlsx', false);

echo "Total de linhas: {$total}\n";
```

## 📖 Documentação Completa dos Métodos

### `read(string $filePath, bool $firstRowAsHeader = false): array`

Lê um arquivo Excel e retorna todos os dados em array.

**Parâmetros:**
- `$filePath`: Caminho completo do arquivo
- `$firstRowAsHeader`: Se TRUE, usa primeira linha como chave do array (array associativo)

**Retorna:** Array com os dados da planilha

```php
// Array numérico
$dados = ForgeExcel::read('arquivo.xlsx');
// [['João', 'joao@email.com'], ['Maria', 'maria@email.com']]

// Array associativo
$dados = ForgeExcel::read('arquivo.xlsx', true);
// [['Nome' => 'João', 'Email' => 'joao@email.com'], ['Nome' => 'Maria', 'Email' => 'maria@email.com']]
```

---

### `write(string $filePath, array $data, string $type = 'xlsx'): bool`

Escreve dados em um arquivo Excel.

**Parâmetros:**
- `$filePath`: Caminho onde o arquivo será salvo
- `$data`: Array de dados (cada item é uma linha)
- `$type`: Tipo do arquivo ('xlsx', 'csv' ou 'ods')

**Retorna:** TRUE se salvou com sucesso

```php
$dados = [
    ['Nome', 'Email'],
    ['João', 'joao@email.com']
];

ForgeExcel::write('arquivo.xlsx', $dados);        // Excel
ForgeExcel::write('arquivo.csv', $dados, 'csv');  // CSV
ForgeExcel::write('arquivo.ods', $dados, 'ods');  // ODS
```

---

### `writeWithSheets(string $filePath, array $sheets): bool`

Escreve dados em múltiplas abas de um arquivo Excel.

**Parâmetros:**
- `$filePath`: Caminho onde o arquivo será salvo
- `$sheets`: Array associativo [nome_aba => dados]

**Retorna:** TRUE se salvou com sucesso

```php
$abas = [
    'Aba1' => [['Coluna1', 'Coluna2'], ['Valor1', 'Valor2']],
    'Aba2' => [['Coluna1', 'Coluna2'], ['Valor1', 'Valor2']]
];

ForgeExcel::writeWithSheets('arquivo.xlsx', $abas);
```

---

### `arrayToExcel(array $associativeArray, bool $includeHeader = true): array`

Converte um array associativo em array formatado para Excel.

**Parâmetros:**
- `$associativeArray`: Array de arrays associativos
- `$includeHeader`: Se TRUE, adiciona linha de cabeçalho automaticamente

**Retorna:** Array formatado para escrita no Excel

```php
$usuarios = [
    ['nome' => 'João', 'email' => 'joao@email.com'],
    ['nome' => 'Maria', 'email' => 'maria@email.com']
];

$dadosExcel = ForgeExcel::arrayToExcel($usuarios);
// [['nome', 'email'], ['João', 'joao@email.com'], ['Maria', 'maria@email.com']]

ForgeExcel::write('usuarios.xlsx', $dadosExcel);
```

---

### `readFirstSheet(string $filePath, bool $firstRowAsHeader = false): array`

Lê apenas a primeira aba de um arquivo Excel (mais rápido).

**Parâmetros:**
- `$filePath`: Caminho do arquivo
- `$firstRowAsHeader`: Se TRUE, usa primeira linha como chave

**Retorna:** Array com dados da primeira aba

```php
$dados = ForgeExcel::readFirstSheet('arquivo.xlsx', true);
```

---

### `readAllSheets(string $filePath, bool $firstRowAsHeader = false): array`

Lê todas as abas de um arquivo Excel separadamente.

**Parâmetros:**
- `$filePath`: Caminho do arquivo
- `$firstRowAsHeader`: Se TRUE, usa primeira linha como chave

**Retorna:** Array associativo [nome_aba => dados]

```php
$todasAbas = ForgeExcel::readAllSheets('arquivo.xlsx', true);

foreach ($todasAbas as $nomeAba => $dados) {
    echo "Processando aba: {$nomeAba}\n";
}
```

---

### `countRows(string $filePath, bool $countHeader = true): int`

Conta quantas linhas tem um arquivo Excel.

**Parâmetros:**
- `$filePath`: Caminho do arquivo
- `$countHeader`: Se FALSE, não conta a primeira linha

**Retorna:** Número total de linhas

```php
$total = ForgeExcel::countRows('arquivo.xlsx');
echo "Total de linhas: {$total}";
```

---

### `readInChunks(string $filePath, int $chunkSize, callable $callback, bool $firstRowAsHeader = false): void`

Processa arquivo Excel em lotes (ideal para arquivos muito grandes).

**Parâmetros:**
- `$filePath`: Caminho do arquivo
- `$chunkSize`: Quantas linhas processar por vez
- `$callback`: Função que recebe cada lote de dados
- `$firstRowAsHeader`: Se TRUE, usa primeira linha como chave

```php
ForgeExcel::readInChunks('arquivo_grande.xlsx', 100, function($lote) {
    // Processa 100 linhas por vez
    foreach ($lote as $linha) {
        processarDados($linha);
    }
});
```

## 🧪 Testando

Execute o arquivo de testes incluído:

```bash
php test.php
```

Este comando vai criar vários arquivos Excel de exemplo demonstrando todos os recursos da biblioteca!

## 🎨 Exemplos de Casos de Uso

### 1. Importar planilha de clientes

```php
$clientes = ForgeExcel::read('clientes.xlsx', true);

foreach ($clientes as $cliente) {
    $pdo->prepare("INSERT INTO clientes (nome, email, telefone) VALUES (?, ?, ?)")
        ->execute([$cliente['Nome'], $cliente['Email'], $cliente['Telefone']]);
}
```

### 2. Exportar relatório mensal

```php
$vendas = $pdo->query("SELECT * FROM vendas WHERE MONTH(data) = 1")->fetchAll(PDO::FETCH_ASSOC);
$dados = ForgeExcel::arrayToExcel($vendas);
ForgeExcel::write('relatorio_janeiro.xlsx', $dados);
```

### 3. Processar arquivo enorme (milhões de linhas)

```php
$total = 0;
$soma = 0;

ForgeExcel::readInChunks('vendas_2023.xlsx', 1000, function($lote) use (&$total, &$soma) {
    foreach ($lote as $venda) {
        $total++;
        $soma += $venda['Valor'];
    }
}, true);

echo "Total de vendas: {$total}\n";
echo "Valor total: R$ {$soma}\n";
```

### 4. Criar dashboard em Excel

```php
$dashboard = [
    'Resumo Geral' => [
        ['Métrica', 'Valor'],
        ['Total de Clientes', 1523],
        ['Total de Vendas', 'R$ 458.230,00'],
        ['Ticket Médio', 'R$ 301,00']
    ],
    'Top 10 Clientes' => ForgeExcel::arrayToExcel($topClientes),
    'Vendas por Mês' => ForgeExcel::arrayToExcel($vendasMes)
];

ForgeExcel::writeWithSheets('dashboard.xlsx', $dashboard);
```

## 🎨 Paleta de Cores

Use cores predefinidas:

```php
$cores = ForgeExcel::colors();
// black, white, red, green, blue, yellow, orange, purple, pink,
// gray, light_gray, dark_gray, cyan, magenta, lime, navy, teal,
// olive, maroon, aqua

$estilo = ForgeExcel::createStyle([
    'color' => $cores['white'],
    'background' => $cores['blue']
]);
```

## 📊 Fórmulas Suportadas

- **Matemáticas**: SUM, AVERAGE, COUNT, MAX, MIN, ROUND, ABS
- **Lógicas**: IF, AND, OR, NOT
- **Condicionais**: COUNTIF, SUMIF, AVERAGEIF
- **Texto**: CONCATENATE, UPPER, LOWER, LEN, LEFT, RIGHT
- **Data**: TODAY, NOW, DATE, YEAR, MONTH, DAY
- E muito mais!

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

## 📝 Licença

Este projeto está sob a licença MIT.

## 👨‍💻 Autor

**Luan Gotardo**
- Email: luan.gotardo.dev@gmail.com

## ⚡ Performance

- ✅ Lê arquivos de 100MB+ sem problemas
- ✅ Usa stream reading (não carrega tudo na memória)
- ✅ Processa milhões de linhas com `readInChunks()`
- ✅ Escrita otimizada e rápida

## 🐛 Problemas Conhecidos

- O box/spout está marcado como "abandoned", mas ainda funciona perfeitamente
- Para arquivos MUITO grandes (1GB+), use sempre `readInChunks()`

## 🆕 Novos Recursos

### Formatação e Estilos

```php
// Criar estilo personalizado
$estilo = ForgeExcel::createStyle([
    'bold' => true,
    'fontSize' => 14,
    'color' => 'FFFFFF',
    'background' => '4472C4',
    'align' => 'center'
]);

// Escrever com formatação
$estilos = [
    0 => ['bold' => true, 'color' => 'FFFFFF', 'background' => '4472C4']
];

ForgeExcel::writeWithStyle('arquivo.xlsx', $dados, $estilos);
```

### Fórmulas do Excel

```php
$dados = [
    ['Produto', 'Quantidade', 'Preço', 'Total'],
    ['Notebook', 2, 3500, '=B2*C2'],
    ['Mouse', 5, 50, '=B3*C3'],
    ['TOTAL', '', '', '=SUM(D2:D3)']
];

ForgeExcel::writeWithFormulas('vendas.xlsx', $dados);
```

### Tabelas com Temas

```php
// 5 temas disponíveis: blue, green, red, orange, purple
ForgeExcel::writeTable('funcionarios.xlsx', $dados, 'blue');
```

### Múltiplas Abas com Estilos

```php
$abas = [
    'Dashboard' => [
        'data' => $dados,
        'headerStyle' => ['bold' => true, 'color' => 'FFFFFF', 'background' => '4472C4']
    ],
    'Detalhes' => [
        'data' => $detalhes,
        'headerStyle' => ['bold' => true, 'color' => 'FFFFFF', 'background' => '70AD47']
    ]
];

ForgeExcel::writeStyledSheets('relatorio.xlsx', $abas);
```

## 📚 Documentação Completa

Para documentação detalhada, consulte:

- **[Guia Completo](docs/GUIA_COMPLETO.md)** - Documentação completa com todos os recursos
- **[Guia de Formatação](docs/FORMATACAO.md)** - Tudo sobre estilos, cores e formatação
- **[Guia de Fórmulas](docs/FORMULAS.md)** - Como usar fórmulas do Excel

## 🧪 Testando

Execute os arquivos de teste incluídos:

```bash
# Testes básicos
php test.php

# Testes avançados (formatação e fórmulas)
php test_advanced.php
```

## 🎯 Roadmap

- [x] Suporte para formatação de células (negrito, cores, etc) ✅
- [x] Suporte para fórmulas ✅
- [x] Tabelas estilizadas com temas ✅
- [ ] Suporte para imagens
- [ ] Suporte para gráficos nativos
- [ ] CLI para conversões rápidas

---

**Feito com ❤️ para facilitar sua vida trabalhando com Excel em PHP!**