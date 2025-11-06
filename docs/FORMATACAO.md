# 🎨 Guia de Formatação e Estilos - ForgeExcel

> **Documentação completa sobre formatação de células, cores, fontes e estilos**

---

## 📑 Índice

1. [Introdução](#introdução)
2. [Conceitos Básicos](#conceitos-básicos)
3. [Criando Estilos](#criando-estilos)
4. [Formatação de Texto](#formatação-de-texto)
5. [Cores e Fundos](#cores-e-fundos)
6. [Alinhamento](#alinhamento)
7. [Bordas](#bordas)
8. [Aplicando Estilos](#aplicando-estilos)
9. [Tabelas Estilizadas](#tabelas-estilizadas)
10. [Exemplos Práticos](#exemplos-práticos)

---

## Introdução

O ForgeExcel oferece recursos completos de formatação para criar planilhas profissionais e visualmente atraentes. Com poucos comandos, você pode aplicar cores, fontes, alinhamentos e muito mais.

### O que você pode fazer:

✅ Aplicar **negrito, itálico e sublinhado**  
✅ Definir **tamanho e tipo de fonte**  
✅ Usar **cores de texto e fundo**  
✅ Configurar **alinhamento de células**  
✅ Adicionar **bordas**  
✅ Criar **tabelas com temas predefinidos**  
✅ Aplicar estilos por **linha ou coluna**  

---

## Conceitos Básicos

### Como Funciona

1. **Criar um estilo** usando `createStyle()`
2. **Aplicar o estilo** ao escrever o arquivo
3. **Usar temas prontos** com `writeTable()`

### Estrutura de um Estilo

```php
$estilo = ForgeExcel::createStyle([
    'bold' => true,           // Negrito
    'italic' => true,         // Itálico
    'underline' => true,      // Sublinhado
    'fontSize' => 14,         // Tamanho
    'fontName' => 'Arial',    // Fonte
    'color' => 'FF0000',      // Cor do texto
    'background' => 'FFFF00', // Cor de fundo
    'align' => 'center',      // Alinhamento
    'wrapText' => true,       // Quebrar texto
    'border' => true          // Bordas
]);
```

---

## Criando Estilos

### Método `createStyle()`

O método principal para criar estilos personalizados.

**Sintaxe:**
```php
ForgeExcel::createStyle(array $options): Style
```

**Retorna:** Objeto `Style` do Box/Spout

### Exemplo Básico

```php
// Estilo simples com negrito
$estiloSimples = ForgeExcel::createStyle(['bold' => true]);

// Estilo complexo
$estiloComplexo = ForgeExcel::createStyle([
    'bold' => true,
    'italic' => true,
    'fontSize' => 16,
    'color' => 'FFFFFF',
    'background' => '4472C4'
]);
```

---

## Formatação de Texto

### Negrito

```php
$estilo = ForgeExcel::createStyle(['bold' => true]);
```

**Exemplo:**
```php
$dados = [
    ['Nome', 'Email'],
    ['João', 'joao@email.com']
];

$estilos = [
    0 => ['bold' => true] // Primeira linha em negrito
];

ForgeExcel::writeWithStyle('arquivo.xlsx', $dados, $estilos);
```

### Itálico

```php
$estilo = ForgeExcel::createStyle(['italic' => true]);
```

### Sublinhado

```php
$estilo = ForgeExcel::createStyle(['underline' => true]);
```

### Combinando Estilos de Texto

```php
$estilo = ForgeExcel::createStyle([
    'bold' => true,
    'italic' => true,
    'underline' => true
]);
```

### Tamanho da Fonte

```php
$estilo = ForgeExcel::createStyle(['fontSize' => 14]);
```

**Tamanhos comuns:**
- `8` - Muito pequeno
- `10` - Pequeno
- `11` - Padrão Excel
- `12` - Médio
- `14` - Grande
- `16` - Muito grande
- `18` - Título

### Nome da Fonte

```php
$estilo = ForgeExcel::createStyle(['fontName' => 'Arial']);
```

**Fontes comuns:**
- `Arial`
- `Calibri` (padrão Excel)
- `Times New Roman`
- `Courier New`
- `Verdana`
- `Tahoma`

### Exemplo Completo de Texto

```php
$estiloTitulo = ForgeExcel::createStyle([
    'bold' => true,
    'fontSize' => 18,
    'fontName' => 'Arial'
]);

$estiloSubtitulo = ForgeExcel::createStyle([
    'italic' => true,
    'fontSize' => 12,
    'fontName' => 'Calibri'
]);

$estiloDestaque = ForgeExcel::createStyle([
    'bold' => true,
    'italic' => true,
    'underline' => true,
    'fontSize' => 14
]);
```

---

## Cores e Fundos

### Formato de Cores

As cores devem ser especificadas em **hexadecimal sem o #**.

**Exemplos:**
- Vermelho: `FF0000`
- Verde: `00FF00`
- Azul: `0000FF`
- Branco: `FFFFFF`
- Preto: `000000`
- Amarelo: `FFFF00`

### Cor do Texto

```php
$estilo = ForgeExcel::createStyle([
    'color' => 'FF0000' // Texto vermelho
]);
```

### Cor de Fundo

```php
$estilo = ForgeExcel::createStyle([
    'background' => 'FFFF00' // Fundo amarelo
]);
```

### Combinando Texto e Fundo

```php
$estilo = ForgeExcel::createStyle([
    'color' => 'FFFFFF',      // Texto branco
    'background' => '4472C4'  // Fundo azul
]);
```

### Paleta de Cores Predefinidas

O ForgeExcel oferece uma paleta com 20 cores prontas:

```php
$cores = ForgeExcel::colors();

// Usar cores predefinidas
$estilo = ForgeExcel::createStyle([
    'color' => $cores['white'],
    'background' => $cores['blue']
]);
```

**Cores disponíveis:**

| Nome | Hex | Visualização |
|------|-----|--------------|
| black | 000000 | ⬛ |
| white | FFFFFF | ⬜ |
| red | FF0000 | 🟥 |
| green | 00FF00 | 🟩 |
| blue | 0000FF | 🟦 |
| yellow | FFFF00 | 🟨 |
| orange | FFA500 | 🟧 |
| purple | 800080 | 🟪 |
| pink | FFC0CB | 💗 |
| gray | 808080 | ▪️ |
| light_gray | D3D3D3 | ▫️ |
| dark_gray | A9A9A9 | ◾ |
| cyan | 00FFFF | 🩵 |
| magenta | FF00FF | 💜 |
| lime | 00FF00 | 💚 |
| navy | 000080 | 💙 |
| teal | 008080 | 💚 |
| olive | 808000 | 💛 |
| maroon | 800000 | ❤️ |
| aqua | 00FFFF | 🩵 |

### Exemplos de Combinações de Cores

**Header Azul Profissional**
```php
$headerAzul = ForgeExcel::createStyle([
    'bold' => true,
    'color' => 'FFFFFF',
    'background' => '4472C4'
]);
```

**Destaque Amarelo**
```php
$destaqueAmarelo = ForgeExcel::createStyle([
    'bold' => true,
    'background' => 'FFFF00'
]);
```

**Erro Vermelho**
```php
$erro = ForgeExcel::createStyle([
    'bold' => true,
    'color' => 'FFFFFF',
    'background' => 'FF0000'
]);
```

**Sucesso Verde**
```php
$sucesso = ForgeExcel::createStyle([
    'bold' => true,
    'color' => 'FFFFFF',
    'background' => '70AD47'
]);
```

**Alerta Laranja**
```php
$alerta = ForgeExcel::createStyle([
    'bold' => true,
    'color' => 'FFFFFF',
    'background' => 'ED7D31'
]);
```

---

## Alinhamento

### Opções de Alinhamento

```php
$alinhamentos = ForgeExcel::alignments();

// Retorna:
[
    'left' => CellAlignment::LEFT,
    'center' => CellAlignment::CENTER,
    'right' => CellAlignment::RIGHT
]
```

### Alinhamento à Esquerda

```php
$estilo = ForgeExcel::createStyle([
    'align' => 'left'
]);
```

### Alinhamento Centralizado

```php
$estilo = ForgeExcel::createStyle([
    'align' => 'center'
]);
```

### Alinhamento à Direita

```php
$estilo = ForgeExcel::createStyle([
    'align' => 'right'
]);
```

### Quebra de Texto

```php
$estilo = ForgeExcel::createStyle([
    'wrapText' => true // Texto longo quebra em múltiplas linhas
]);
```

### Exemplo Prático de Alinhamento

```php
$dados = [
    ['Nome', 'Quantidade', 'Valor', 'Descrição'],
    ['Produto A', 100, 2500.00, 'Descrição muito longa que precisa quebrar'],
    ['Produto B', 50, 1250.00, 'Outra descrição longa']
];

$estilosColunas = [
    0 => ['align' => 'left'],           // Nome à esquerda
    1 => ['align' => 'center'],         // Quantidade centralizado
    2 => ['align' => 'right'],          // Valor à direita
    3 => ['align' => 'left', 'wrapText' => true] // Descrição com quebra
];

ForgeExcel::writeWithStyle('produtos.xlsx', $dados, [], $estilosColunas);
```

---

## Bordas

### Adicionar Bordas Simples

```php
$estilo = ForgeExcel::createStyle([
    'border' => true
]);
```

Isso adiciona bordas finas em todos os lados da célula.

### Estilos de Borda Disponíveis

O Box/Spout suporta diferentes estilos de borda:

- `Border::STYLE_THIN` - Linha fina (padrão)
- `Border::STYLE_MEDIUM` - Linha média
- `Border::STYLE_THICK` - Linha grossa
- `Border::STYLE_DASHED` - Linha tracejada
- `Border::STYLE_DOTTED` - Linha pontilhada
- `Border::STYLE_DOUBLE` - Linha dupla

### Personalizar Bordas

```php
use Box\Spout\Common\Entity\Style\Border;

$estilo = ForgeExcel::createStyle([
    'border' => true,
    'borderStyle' => Border::STYLE_MEDIUM,
    'borderColor' => 'FF0000' // Borda vermelha
]);
```

---

## Aplicando Estilos

### Método 1: Estilos por Linha

```php
$dados = [
    ['Nome', 'Email'],           // Linha 0
    ['João', 'joao@email.com'],  // Linha 1
    ['Maria', 'maria@email.com'] // Linha 2
];

$estilosPorLinha = [
    0 => ['bold' => true, 'background' => '4472C4', 'color' => 'FFFFFF'], // Header
    1 => ['background' => 'D9E1F2'], // Linha 1 com fundo azul claro
    2 => ['background' => 'FFFFFF']  // Linha 2 com fundo branco
];

ForgeExcel::writeWithStyle('arquivo.xlsx', $dados, $estilosPorLinha);
```

### Método 2: Estilos por Coluna

```php
$dados = [
    ['Nome', 'Idade', 'Salário'],
    ['João', 25, 5000.00],
    ['Maria', 30, 6000.00]
];

$estilosPorColuna = [
    0 => ['align' => 'left'],   // Coluna 0 (Nome)
    1 => ['align' => 'center'], // Coluna 1 (Idade)
    2 => ['align' => 'right', 'bold' => true] // Coluna 2 (Salário)
];

ForgeExcel::writeWithStyle('arquivo.xlsx', $dados, [], $estilosPorColuna);
```

### Método 3: Combinando Linha e Coluna

```php
$dados = [
    ['Produto', 'Quantidade', 'Valor'],
    ['Notebook', 10, 35000.00],
    ['Mouse', 50, 2500.00]
];

// Estilos por linha (prioridade maior)
$estilosPorLinha = [
    0 => ['bold' => true, 'color' => 'FFFFFF', 'background' => '4472C4']
];

// Estilos por coluna (aplicado quando não há estilo de linha)
$estilosPorColuna = [
    1 => ['align' => 'center'],
    2 => ['align' => 'right', 'bold' => true]
];

ForgeExcel::writeWithStyle('arquivo.xlsx', $dados, $estilosPorLinha, $estilosPorColuna);
```

---

## Tabelas Estilizadas

### Método `writeTable()`

Cria tabelas com temas profissionais predefinidos.

**Sintaxe:**
```php
ForgeExcel::writeTable(string $filePath, array $data, string $theme = 'blue'): bool
```

### Temas Disponíveis

#### 1. Blue (Azul Profissional)

```php
ForgeExcel::writeTable('tabela.xlsx', $dados, 'blue');
```

- **Header:** Azul escuro (#4472C4) com texto branco e negrito
- **Linhas ímpares:** Azul claro (#D9E1F2)
- **Linhas pares:** Branco (#FFFFFF)

#### 2. Green (Verde Natureza)

```php
ForgeExcel::writeTable('tabela.xlsx', $dados, 'green');
```

- **Header:** Verde (#70AD47) com texto branco e negrito
- **Linhas ímpares:** Verde claro (#E2EFDA)
- **Linhas pares:** Branco (#FFFFFF)

#### 3. Red (Vermelho Corporativo)

```php
ForgeExcel::writeTable('tabela.xlsx', $dados, 'red');
```

- **Header:** Vermelho alaranjado (#C55A11) com texto branco e negrito
- **Linhas ímpares:** Laranja claro (#FCE4D6)
- **Linhas pares:** Branco (#FFFFFF)

#### 4. Orange (Laranja Vibrante)

```php
ForgeExcel::writeTable('tabela.xlsx', $dados, 'orange');
```

- **Header:** Laranja (#ED7D31) com texto branco e negrito
- **Linhas ímpares:** Bege (#FBE5D6)
- **Linhas pares:** Branco (#FFFFFF)

#### 5. Purple (Roxo Elegante)

```php
ForgeExcel::writeTable('tabela.xlsx', $dados, 'purple');
```

- **Header:** Roxo escuro (#7030A0) com texto branco e negrito
- **Linhas ímpares:** Lilás claro (#E4DFEC)
- **Linhas pares:** Branco (#FFFFFF)

### Exemplo Completo de Tabela

```php
$funcionarios = [
    ['ID', 'Nome', 'Departamento', 'Salário', 'Admissão'],
    [1, 'João Silva', 'TI', 8500.00, '2020-03-15'],
    [2, 'Maria Santos', 'RH', 12000.00, '2018-06-01'],
    [3, 'Pedro Costa', 'TI', 15000.00, '2019-01-20'],
    [4, 'Ana Oliveira', 'Marketing', 9500.00, '2021-02-10']
];

// Cria tabela com tema azul
ForgeExcel::writeTable('funcionarios.xlsx', $funcionarios, 'blue');
```

---

## Exemplos Práticos

### Exemplo 1: Relatório Financeiro

```php
$dados = [
    ['Descrição', 'Valor', 'Status'],
    ['Receitas', 150000.00, 'Recebido'],
    ['Despesas', 98000.00, 'Pago'],
    ['Lucro', 52000.00, 'Positivo']
];

$cores = ForgeExcel::colors();

$estilos = [
    0 => [ // Header
        'bold' => true,
        'fontSize' => 13,
        'color' => $cores['white'],
        'background' => '203864',
        'align' => 'center'
    ],
    1 => ['color' => $cores['green'], 'bold' => true], // Receitas em verde
    2 => ['color' => $cores['red'], 'bold' => true],   // Despesas em vermelho
    3 => ['color' => $cores['blue'], 'bold' => true, 'fontSize' => 12] // Lucro em azul
];

$estilosColunas = [
    1 => ['align' => 'right'], // Valor alinhado à direita
    2 => ['align' => 'center'] // Status centralizado
];

ForgeExcel::writeWithStyle('relatorio_financeiro.xlsx', $dados, $estilos, $estilosColunas);
```

### Exemplo 2: Dashboard de Vendas

```php
$cores = ForgeExcel::colors();

$dados = [
    ['Métrica', 'Valor', 'Meta', 'Status'],
    ['Vendas Totais', 'R$ 375.000', 'R$ 350.000', '✓'],
    ['Novos Clientes', '127', '100', '✓'],
    ['Taxa de Conversão', '15,3%', '12,0%', '✓'],
    ['Ticket Médio', 'R$ 2.952', 'R$ 3.000', '✗']
];

$estilos = [
    0 => [ // Header
        'bold' => true,
        'fontSize' => 14,
        'color' => $cores['white'],
        'background' => '4472C4',
        'align' => 'center'
    ],
    1 => ['background' => 'E7F3E7'], // Verde claro
    2 => ['background' => 'E7F3E7'],
    3 => ['background' => 'E7F3E7'],
    4 => ['background' => 'FFE7E7']  // Vermelho claro
];

ForgeExcel::writeWithStyle('dashboard.xlsx', $dados, $estilos);
```

### Exemplo 3: Lista de Tarefas

```php
$tarefas = [
    ['ID', 'Tarefa', 'Responsável', 'Prazo', 'Status'],
    [1, 'Desenvolver API', 'João', '2024-02-15', 'Em Andamento'],
    [2, 'Criar Documentação', 'Maria', '2024-02-20', 'Pendente'],
    [3, 'Fazer Testes', 'Pedro', '2024-02-10', 'Concluído'],
    [4, 'Deploy Produção', 'Ana', '2024-02-25', 'Pendente']
];

ForgeExcel::writeTable('tarefas.xlsx', $tarefas, 'purple');
```

### Exemplo 4: Planilha de Notas

```php
$cores = ForgeExcel::colors();

$notas = [
    ['Aluno', 'P1', 'P2', 'P3', 'Média', 'Situação'],
    ['João Silva', 8.5, 7.0, 9.0, 8.17, 'Aprovado'],
    ['Maria Santos', 9.5, 9.0, 8.5, 9.00, 'Aprovado'],
    ['Pedro Costa', 6.0, 5.5, 6.5, 6.00, 'Reprovado'],
    ['Ana Oliveira', 7.5, 8.0, 7.0, 7.50, 'Aprovado']
];

$estilos = [
    0 => [ // Header
        'bold' => true,
        'color' => $cores['white'],
        'background' => '70AD47',
        'align' => 'center'
    ],
    3 => ['color' => $cores['red'], 'bold' => true] // Reprovado em vermelho
];

$estilosColunas = [
    1 => ['align' => 'center'], // Notas centralizadas
    2 => ['align' => 'center'],
    3 => ['align' => 'center'],
    4 => ['align' => 'center', 'bold' => true],
    5 => ['align' => 'center', 'bold' => true]
];

ForgeExcel::writeWithStyle('notas.xlsx', $notas, $estilos, $estilosColunas);
```

### Exemplo 5: Calendário de Eventos

```php
$eventos = [
    ['Data', 'Evento', 'Local', 'Participantes', 'Tipo'],
    ['15/02/2024', 'Reunião de Equipe', 'Sala 1', 12, 'Interno'],
    ['18/02/2024', 'Workshop PHP', 'Auditório', 45, 'Treinamento'],
    ['22/02/2024', 'Sprint Planning', 'Sala 2', 8, 'Interno'],
    ['25/02/2024', 'Apresentação Cliente', 'Online', 5, 'Externo']
];

// Usar tema laranja para eventos
ForgeExcel::writeTable('calendario.xlsx', $eventos, 'orange');
```

---

## Dicas e Truques

### 1. Reutilize Estilos

```php
class MinhasPaletas {
    public static function headerPadrao() {
        return [
            'bold' => true,
            'fontSize' => 12,
            'color' => 'FFFFFF',
            'background' => '4472C4'
        ];
    }
    
    public static function totalDestaque() {
        return [
            'bold' => true,
            'fontSize' => 11,
            'background' => 'FFFF00'
        ];
    }
    
    public static function valorMonetario() {
        return [
            'align' => 'right',
            'fontName' => 'Courier New'
        ];
    }
}

// Uso
$estilos = [
    0 => MinhasPaletas::headerPadrao(),
    10 => MinhasPaletas::totalDestaque()
];
```

### 2. Cores Corporativas

```php
class CoresCorporativas {
    const PRIMARIA = '4472C4';
    const SECUNDARIA = '70AD47';
    const DESTAQUE = 'FFA500';
    const ERRO = 'FF0000';
    const SUCESSO = '00FF00';
    const ALERTA = 'FFFF00';
}

$estilo = ForgeExcel::createStyle([
    'color' => 'FFFFFF',
    'background' => CoresCorporativas::PRIMARIA
]);
```

### 3. Formatação Condicional Manual

```php
$dados = [
    ['Produto', 'Estoque', 'Status'],
    ['Produto A', 5, 'Baixo'],
    ['Produto B', 50, 'Normal'],
    ['Produto C', 2, 'Crítico']
];

$estilos = [];
foreach ($dados as $index => $linha) {
    if ($index === 0) continue; // Pula header
    
    $estoque = $linha[1];
    
    if ($estoque < 5) {
        $estilos[$index] = ['color' => 'FFFFFF', 'background' => 'FF0000']; // Vermelho
    } elseif ($estoque < 20) {
        $estilos[$index] = ['color' => '000000', 'background' => 'FFFF00']; // Amarelo
    } else {
        $estilos[$index] = ['color' => 'FFFFFF', 'background' => '00FF00']; // Verde
    }
}

$estilos[0] = ['bold' => true]; // Header

ForgeExcel::writeWithStyle('estoque.xlsx', $dados, $estilos);
```

---

## Conclusão

Com esses recursos de formatação, você pode criar planilhas Excel profissionais e visualmente atraentes. Experimente combinar diferentes estilos para encontrar o visual perfeito para seus relatórios!

**Próximos passos:**
- Explore o [Guia de Fórmulas](FORMULAS.md)
- Veja exemplos no [Guia Completo](GUIA_COMPLETO.md)
- Execute os testes: `php test_advanced.php`

---

**Desenvolvido com ❤️ por Luan Gotardo**