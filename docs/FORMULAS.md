# 📊 Guia de Fórmulas - ForgeExcel

> **Documentação completa sobre criação e uso de fórmulas do Excel**

---

## 📑 Índice

1. [Introdução](#introdução)
2. [Conceitos Básicos](#conceitos-básicos)
3. [Fórmulas Matemáticas](#fórmulas-matemáticas)
4. [Fórmulas Estatísticas](#fórmulas-estatísticas)
5. [Fórmulas Lógicas](#fórmulas-lógicas)
6. [Fórmulas de Texto](#fórmulas-de-texto)
7. [Fórmulas de Data](#fórmulas-de-data)
8. [Referências de Células](#referências-de-células)
9. [Fórmulas Avançadas](#fórmulas-avançadas)
10. [Exemplos Práticos](#exemplos-práticos)

---

## Introdução

O ForgeExcel permite criar planilhas com fórmulas do Excel que são calculadas automaticamente quando o arquivo é aberto. Isso é perfeito para:

✅ **Relatórios dinâmicos** que atualizam automaticamente  
✅ **Cálculos financeiros** complexos  
✅ **Análises estatísticas**  
✅ **Validações e verificações** automáticas  
✅ **Dashboards interativos**  

---

## Conceitos Básicos

### Como Escrever Fórmulas

No ForgeExcel, fórmulas são strings que começam com o sinal `=`:

```php
$dados = [
    ['A', 'B', 'Total'],
    [10, 20, '=A2+B2']  // Fórmula: soma A2 + B2
];
```

### Método Principal

```php
ForgeExcel::writeWithFormulas(string $filePath, array $data, array $headerStyle = []): bool
```

**Exemplo básico:**
```php
$dados = [
    ['Produto', 'Quantidade', 'Preço', 'Total'],
    ['Notebook', 2, 3500, '=B2*C2']
];

ForgeExcel::writeWithFormulas('vendas.xlsx', $dados);
```

### Sintaxe de Células

| Referência | Significado | Exemplo |
|------------|-------------|---------|
| `A1` | Célula A1 | `=A1*2` |
| `B2:B10` | Intervalo B2 até B10 | `=SUM(B2:B10)` |
| `$A$1` | Referência absoluta | `=$A$1*B2` |
| `$A1` | Coluna fixa, linha relativa | `=$A1*2` |
| `A$1` | Linha fixa, coluna relativa | `=A$1*2` |

---

## Fórmulas Matemáticas

### Operações Básicas

#### Adição (+)
```php
$dados = [
    ['A', 'B', 'Soma'],
    [10, 20, '=A2+B2']  // Resultado: 30
];
```

#### Subtração (-)
```php
$dados = [
    ['Receita', 'Despesa', 'Lucro'],
    [50000, 35000, '=A2-B2']  // Resultado: 15000
];
```

#### Multiplicação (*)
```php
$dados = [
    ['Quantidade', 'Preço', 'Total'],
    [5, 100, '=A2*B2']  // Resultado: 500
];
```

#### Divisão (/)
```php
$dados = [
    ['Total', 'Quantidade', 'Média'],
    [1000, 4, '=A2/B2']  // Resultado: 250
];
```

#### Exponenciação (^)
```php
$dados = [
    ['Base', 'Expoente', 'Resultado'],
    [2, 8, '=A2^B2']  // Resultado: 256
];
```

### SUM - Soma

Soma valores de um intervalo.

**Sintaxe:** `=SUM(intervalo)`

```php
$dados = [
    ['Mês', 'Valor'],
    ['Janeiro', 1000],
    ['Fevereiro', 1500],
    ['Março', 1200],
    ['', ''],
    ['TOTAL', '=SUM(B2:B4)']  // Resultado: 3700
];
```

**Com múltiplos intervalos:**
```php
['Total Geral', '=SUM(B2:B5,D2:D5,F2:F5)']
```

### SUMIF - Soma Condicional

Soma valores que atendem uma condição.

**Sintaxe:** `=SUMIF(intervalo_teste, critério, intervalo_soma)`

```php
$dados = [
    ['Produto', 'Categoria', 'Valor'],
    ['Item A', 'Eletrônicos', 1000],
    ['Item B', 'Móveis', 500],
    ['Item C', 'Eletrônicos', 1500],
    ['', '', ''],
    ['Total Eletrônicos', '', '=SUMIF(B2:B4,"Eletrônicos",C2:C4)']  // 2500
];
```

### PRODUCT - Multiplicação

Multiplica valores de um intervalo.

**Sintaxe:** `=PRODUCT(intervalo)`

```php
$dados = [
    ['Fator', 'Valor'],
    ['Fator 1', 2],
    ['Fator 2', 3],
    ['Fator 3', 4],
    ['Produto', '=PRODUCT(B2:B4)']  // Resultado: 24
];
```

### ROUND - Arredondamento

Arredonda um número.

**Sintaxe:** `=ROUND(número, decimais)`

```php
$dados = [
    ['Valor', 'Arredondado'],
    [15.678, '=ROUND(A2,2)'],  // 15.68
    [23.234, '=ROUND(A3,1)'],  // 23.2
    [7.5, '=ROUND(A4,0)']      // 8
];
```

### ABS - Valor Absoluto

Retorna o valor absoluto (sem sinal).

**Sintaxe:** `=ABS(número)`

```php
$dados = [
    ['Valor', 'Absoluto'],
    [-50, '=ABS(A2)'],   // 50
    [30, '=ABS(A3)'],    // 30
    [-100, '=ABS(A4)']   // 100
];
```

### MOD - Resto da Divisão

Retorna o resto de uma divisão.

**Sintaxe:** `=MOD(dividendo, divisor)`

```php
$dados = [
    ['Número', 'Resto por 3'],
    [10, '=MOD(A2,3)'],  // 1
    [15, '=MOD(A3,3)'],  // 0
    [7, '=MOD(A4,3)']    // 1
];
```

---

## Fórmulas Estatísticas

### AVERAGE - Média

Calcula a média aritmética.

**Sintaxe:** `=AVERAGE(intervalo)`

```php
$dados = [
    ['Valor'],
    [100],
    [150],
    [200],
    [175],
    [''],
    ['Média', '=AVERAGE(A2:A5)']  // 156.25
];
```

### COUNT - Contar Números

Conta quantas células contêm números.

**Sintaxe:** `=COUNT(intervalo)`

```php
$dados = [
    ['Valor'],
    [100],
    ['Texto'],
    [200],
    [300],
    [''],
    ['Quantidade', '=COUNT(A2:A5)']  // 3
];
```

### COUNTA - Contar Não Vazias

Conta células não vazias.

**Sintaxe:** `=COUNTA(intervalo)`

```php
['Total Preenchido', '=COUNTA(A2:A10)']
```

### COUNTIF - Contar com Condição

Conta células que atendem critério.

**Sintaxe:** `=COUNTIF(intervalo, critério)`

```php
$dados = [
    ['Aluno', 'Situação'],
    ['João', 'Aprovado'],
    ['Maria', 'Reprovado'],
    ['Pedro', 'Aprovado'],
    ['Ana', 'Aprovado'],
    ['', ''],
    ['Aprovados', '=COUNTIF(B2:B5,"Aprovado")'],    // 3
    ['Reprovados', '=COUNTIF(B2:B5,"Reprovado")']   // 1
];
```

### MAX - Valor Máximo

Retorna o maior valor.

**Sintaxe:** `=MAX(intervalo)`

```php
$dados = [
    ['Valor'],
    [100],
    [250],
    [150],
    [300],
    [''],
    ['Máximo', '=MAX(A2:A5)']  // 300
];
```

### MIN - Valor Mínimo

Retorna o menor valor.

**Sintaxe:** `=MIN(intervalo)`

```php
['Mínimo', '=MIN(A2:A10)']
```

### MEDIAN - Mediana

Retorna o valor do meio.

**Sintaxe:** `=MEDIAN(intervalo)`

```php
['Mediana', '=MEDIAN(A2:A10)']
```

### MODE - Moda

Retorna o valor mais frequente.

**Sintaxe:** `=MODE(intervalo)`

```php
['Moda', '=MODE(A2:A10)']
```

---

## Fórmulas Lógicas

### IF - Condicional

Executa teste lógico.

**Sintaxe:** `=IF(teste, se_verdadeiro, se_falso)`

```php
$dados = [
    ['Aluno', 'Nota', 'Situação'],
    ['João', 8.5, '=IF(B2>=7,"Aprovado","Reprovado")'],
    ['Maria', 6.0, '=IF(B3>=7,"Aprovado","Reprovado")'],
    ['Pedro', 7.5, '=IF(B4>=7,"Aprovado","Reprovado")']
];
```

**IF aninhado:**
```php
[
    'Status',
    '=IF(A2>=9,"Excelente",IF(A2>=7,"Bom",IF(A2>=5,"Regular","Insuficiente")))'
]
```

### AND - E Lógico

Retorna TRUE se todas condições forem verdadeiras.

**Sintaxe:** `=AND(condição1, condição2, ...)`

```php
$dados = [
    ['Nome', 'Nota1', 'Nota2', 'Aprovado'],
    ['João', 7.5, 8.0, '=IF(AND(B2>=7,C2>=7),"Sim","Não")']
];
```

### OR - OU Lógico

Retorna TRUE se pelo menos uma condição for verdadeira.

**Sintaxe:** `=OR(condição1, condição2, ...)`

```php
[
    'Desconto',
    '=IF(OR(A2>1000,B2="VIP"),"Sim","Não")'
]
```

### NOT - Negação

Inverte o resultado lógico.

**Sintaxe:** `=NOT(lógico)`

```php
['Inativo', '=NOT(A2="Ativo")']
```

---

## Fórmulas de Texto

### CONCATENATE - Concatenar

Junta textos.

**Sintaxe:** `=CONCATENATE(texto1, texto2, ...)`

```php
$dados = [
    ['Nome', 'Sobrenome', 'Nome Completo'],
    ['João', 'Silva', '=CONCATENATE(A2," ",B2)']  // João Silva
];
```

**Operador alternativo (&):**
```php
['Nome Completo', '=A2&" "&B2']
```

### UPPER - Maiúsculas

Converte para maiúsculas.

**Sintaxe:** `=UPPER(texto)`

```php
['Maiúsculas', '=UPPER(A2)']
```

### LOWER - Minúsculas

Converte para minúsculas.

**Sintaxe:** `=LOWER(texto)`

```php
['Minúsculas', '=LOWER(A2)']
```

### PROPER - Primeira Letra Maiúscula

Capitaliza cada palavra.

**Sintaxe:** `=PROPER(texto)`

```php
['Capitalizado', '=PROPER(A2)']
```

### LEN - Comprimento

Retorna o número de caracteres.

**Sintaxe:** `=LEN(texto)`

```php
['Tamanho', '=LEN(A2)']
```

### LEFT - Primeiros Caracteres

Extrai caracteres da esquerda.

**Sintaxe:** `=LEFT(texto, quantidade)`

```php
['Iniciais', '=LEFT(A2,3)']
```

### RIGHT - Últimos Caracteres

Extrai caracteres da direita.

**Sintaxe:** `=RIGHT(texto, quantidade)`

```php
['Finais', '=RIGHT(A2,3)']
```

### MID - Caracteres do Meio

Extrai caracteres do meio.

**Sintaxe:** `=MID(texto, início, quantidade)`

```php
['Meio', '=MID(A2,3,5)']
```

---

## Fórmulas de Data

### TODAY - Data Atual

Retorna a data atual.

**Sintaxe:** `=TODAY()`

```php
['Data Atual', '=TODAY()']
```

### NOW - Data e Hora Atual

Retorna data e hora atual.

**Sintaxe:** `=NOW()`

```php
['Timestamp', '=NOW()']
```

### DATE - Criar Data

Cria uma data a partir de ano, mês, dia.

**Sintaxe:** `=DATE(ano, mês, dia)`

```php
['Data', '=DATE(2024,12,25)']
```

### YEAR, MONTH, DAY - Extrair Data

Extrai partes de uma data.

```php
$dados = [
    ['Data', 'Ano', 'Mês', 'Dia'],
    ['2024-01-15', '=YEAR(A2)', '=MONTH(A2)', '=DAY(A2)']
];
```

### DATEDIF - Diferença de Datas

Calcula diferença entre datas.

**Sintaxe:** `=DATEDIF(data_inicial, data_final, unidade)`

Unidades:
- "D" - Dias
- "M" - Meses
- "Y" - Anos

```php
['Dias', '=DATEDIF(A2,B2,"D")']
```

---

## Referências de Células

### Referência Relativa

Move-se quando copiada.

```php
$dados = [
    ['A', 'B', 'Soma'],
    [10, 20, '=A2+B2'],  // Na linha 2
    [30, 40, '=A3+B3']   // Na linha 3 (ajustou automaticamente)
];
```

### Referência Absoluta

Não muda quando copiada.

```php
$dados = [
    ['Preço Base', 1000],
    ['', ''],
    ['Item', 'Quantidade', 'Total'],
    ['Item 1', 2, '=B4*$B$1'],  // Sempre usa B1
    ['Item 2', 3, '=B5*$B$1'],  // Sempre usa B1
    ['Item 3', 5, '=B6*$B$1']   // Sempre usa B1
];
```

### Referência Mista

Parte fixa, parte relativa.

```php
// Coluna fixa, linha relativa
['Total', '=$A2*B2']

// Linha fixa, coluna relativa
['Total', '=A$1*B2']
```

### Exemplo Completo - Tabela de Multiplicação

```php
$dados = [
    ['X', 1, 2, 3, 4, 5],
    [1, '=$A2*B$1', '=$A2*C$1', '=$A2*D$1', '=$A2*E$1', '=$A2*F$1'],
    [2, '=$A3*B$1', '=$A3*C$1', '=$A3*D$1', '=$A3*E$1', '=$A3*F$1'],
    [3, '=$A4*B$1', '=$A4*C$1', '=$A4*D$1', '=$A4*E$1', '=$A4*F$1'],
    [4, '=$A5*B$1', '=$A5*C$1', '=$A5*D$1', '=$A5*E$1', '=$A5*F$1'],
    [5, '=$A6*B$1', '=$A6*C$1', '=$A6*D$1', '=$A6*E$1', '=$A6*F$1']
];

ForgeExcel::writeWithFormulas('tabuada.xlsx', $dados);
```

---

## Fórmulas Avançadas

### VLOOKUP - Procura Vertical

Procura valor em tabela.

**Sintaxe:** `=VLOOKUP(valor_procurado, tabela, coluna, [correspondência_exata])`

```php
// Requer configuração manual no Excel após criação
['Preço', '=VLOOKUP(A2,Produtos!A:B,2,FALSE)']
```

### SUMIFS - Soma com Múltiplas Condições

**Sintaxe:** `=SUMIFS(intervalo_soma, intervalo_critério1, critério1, ...)`

```php
[
    'Total',
    '=SUMIFS(C2:C100,A2:A100,"Produto A",B2:B100,">1000")'
]
```

### AVERAGEIF - Média Condicional

**Sintaxe:** `=AVERAGEIF(intervalo_critério, critério, intervalo_média)`

```php
['Média Aprovados', '=AVERAGEIF(C2:C10,"Aprovado",B2:B10)']
```

### IFERROR - Tratar Erros

Executa alternativa se houver erro.

**Sintaxe:** `=IFERROR(fórmula, valor_se_erro)`

```php
['Resultado', '=IFERROR(A2/B2,"Divisão inválida")']
```

---

## Exemplos Práticos

### Exemplo 1: Relatório Financeiro Completo

```php
$dados = [
    ['Mês', 'Receita', 'Despesas', 'Lucro', 'Margem %'],
    ['Janeiro', 50000, 35000, '=B2-C2', '=(D2/B2)*100'],
    ['Fevereiro', 62000, 42000, '=B3-C3', '=(D3/B3)*100'],
    ['Março', 58000, 38000, '=B4-C4', '=(D4/B4)*100'],
    ['Abril', 71000, 48000, '=B5-C5', '=(D5/B5)*100'],
    ['Maio', 65000, 44000, '=B6-C6', '=(D6/B6)*100'],
    ['Junho', 69000, 46000, '=B7-C7', '=(D7/B7)*100'],
    ['', '', '', '', ''],
    ['TOTAIS', '=SUM(B2:B7)', '=SUM(C2:C7)', '=SUM(D2:D7)', ''],
    ['MÉDIAS', '=AVERAGE(B2:B7)', '=AVERAGE(C2:C7)', '=AVERAGE(D2:D7)', '=AVERAGE(E2:E7)'],
    ['MÁXIMO', '=MAX(B2:B7)', '=MAX(C2:C7)', '=MAX(D2:D7)', '=MAX(E2:E7)'],
    ['MÍNIMO', '=MIN(B2:B7)', '=MIN(C2:C7)', '=MIN(D2:D7)', '=MIN(E2:E7)']
];

$headerStyle = [
    'bold' => true,
    'color' => 'FFFFFF',
    'background' => '203864',
    'fontSize' => 11
];

ForgeExcel::writeWithFormulas('financeiro.xlsx', $dados, $headerStyle);
```

### Exemplo 2: Controle de Estoque com Alertas

```php
$dados = [
    ['Produto', 'Estoque Atual', 'Estoque Mínimo', 'Reposição', 'Status'],
    ['Notebook', 5, 10, '=IF(B2<C2,C2-B2,0)', '=IF(B2<C2,"REPOR","OK")'],
    ['Mouse', 50, 20, '=IF(B3<C3,C3-B3,0)', '=IF(B3<C3,"REPOR","OK")'],
    ['Teclado', 15, 15, '=IF(B4<C4,C4-B4,0)', '=IF(B4<C4,"REPOR","OK")'],
    ['Monitor', 3, 8, '=IF(B5<C5,C5-B5,0)', '=IF(B5<C5,"REPOR","OK")'],
    ['Webcam', 25, 10, '=IF(B6<C6,C6-B6,0)', '=IF(B6<C6,"REPOR","OK")'],
    ['', '', '', '', ''],
    ['Total a Repor', '', '', '=SUM(D2:D6)', '']
];

ForgeExcel::writeWithFormulas('estoque.xlsx', $dados);
```

### Exemplo 3: Folha de Pagamento

```php
$dados = [
    ['Nome', 'Sal. Base', 'H.Extra', 'Vlr H.Extra', 'Total Extras', 'Bruto', 'INSS 11%', 'IRRF 15%', 'Líquido'],
    ['João', 3000, 10, 25, '=C2*D2', '=B2+E2', '=F2*0.11', '=F2*0.15', '=F2-G2-H2'],
    ['Maria', 4500, 5, 37.50, '=C3*D3', '=B3+E3', '=F3*0.11', '=F3*0.15', '=F3-G3-H3'],
    ['Pedro', 5000, 8, 41.67, '=C4*D4', '=B4+E4', '=F4*0.11', '=F4*0.15', '=F4-G4-H4'],
    ['Ana', 3500, 12, 29.17, '=C5*D5', '=B5+E5', '=F5*0.11', '=F5*0.15', '=F5-G5-H5'],
    ['', '', '', '', '', '', '', '', ''],
    ['TOTAIS', '=SUM(B2:B5)', '=SUM(C2:C5)', '', '=SUM(E2:E5)', '=SUM(F2:F5)', '=SUM(G2:G5)', '=SUM(H2:H5)', '=SUM(I2:I5)']
];

ForgeExcel::writeWithFormulas('folha_pagamento.xlsx', $dados);
```

### Exemplo 4: Análise de Vendas por Região

```php
$dados = [
    ['Região', 'Q1', 'Q2', 'Q3', 'Q4', 'Total Anual', 'Média', '% do Total'],
    ['Norte', 120000, 135000, 145000, 150000, '=SUM(B2:E2)', '=AVERAGE(B2:E2)', '=F2/$F$7*100'],
    ['Sul', 150000, 165000, 170000, 180000, '=SUM(B3:E3)', '=AVERAGE(B3:E3)', '=F3/$F$7*100'],
    ['Leste', 100000, 110000, 120000, 125000, '=SUM(B4:E4)', '=AVERAGE(B4:E4)', '=F4/$F$7*100'],
    ['Oeste', 130000, 140000, 155000, 160000, '=SUM(B5:E5)', '=AVERAGE(B5:E5)', '=F5/$F$7*100'],
    ['Centro', 90000, 95000, 100000, 105000, '=SUM(B6:E6)', '=AVERAGE(B6:E6)', '=F6/$F$7*100'],
    ['', '', '', '', '', '', '', ''],
    ['TOTAL', '=SUM(B2:B6)', '=SUM(C2:C6)', '=SUM(D2:D6)', '=SUM(E2:E6)', '=SUM(F2:F6)', '=AVERAGE(G2:G6)', '100%']
];

ForgeExcel::writeWithFormulas('vendas_regioes.xlsx', $dados);
```

### Exemplo 5: Controle de Notas Escolares

```php
$dados = [
    ['Aluno', 'Prova 1', 'Prova 2', 'Prova 3', 'Trabalho', 'Média', 'Situação', 'Falta p/ 7'],
    ['João', 8.5, 7.0, 9.0, 8.0, '=AVERAGE(B2:E2)', '=IF(F2>=7,"Aprovado","Reprovado")', '=IF(F2<7,7-F2,"")'],
    ['Maria', 9.5, 9.0, 8.5, 9.5, '=AVERAGE(B3:E3)', '=IF(F3>=7,"Aprovado","Reprovado")', '=IF(F3<7,7-F3,"")'],
    ['Pedro', 6.0, 5.5, 6.5, 7.0, '=AVERAGE(B4:E4)', '=IF(F4>=7,"Aprovado","Reprovado")', '=IF(F4<7,7-F4,"")'],
    ['Ana', 7.5, 8.0, 7.0, 8.5, '=AVERAGE(B5:E5)', '=IF(F5>=7,"Aprovado","Reprovado")', '=IF(F5<7,7-F5,"")'],
    ['Carlos', 5.0, 6.0, 5.5, 6.0, '=AVERAGE(B6:E6)', '=IF(F6>=7,"Aprovado","Reprovado")', '=IF(F6<7,7-F6,"")'],
    ['', '', '', '', '', '', '', ''],
    ['Média Turma', '=AVERAGE(B2:B6)', '=AVERAGE(C2:C6)', '=AVERAGE(D2:D6)', '=AVERAGE(E2:E6)', '=AVERAGE(F2:F6)', '', ''],
    ['Aprovados', '', '', '', '', '', '=COUNTIF(G2:G6,"Aprovado")', ''],
    ['Reprovados', '', '', '', '', '', '=COUNTIF(G2:G6,"Reprovado")', ''],
    ['Taxa Aprovação', '', '', '', '', '', '=I8/(I8+I9)*100&"%"', '']
];

ForgeExcel::writeWithFormulas('notas_escolares.xlsx', $dados);
```

### Exemplo 6: Cálculo de Impostos

```php
$dados = [
    ['Produto', 'Valor Base', 'ICMS 18%', 'IPI 10%', 'PIS 1.65%', 'COFINS 7.6%', 'Valor Final'],
    ['Produto A', 1000, '=B2*0.18', '=B2*0.10', '=B2*0.0165', '=B2*0.076', '=B2+C2+D2+E2+F2'],
    ['Produto B', 2500, '=B3*0.18', '=B3*0.10', '=B3*0.0165', '=B3*0.076', '=B3+C3+D3+E3+F3'],
    ['Produto C', 5000, '=B4*0.18', '=B4*0.10', '=B4*0.0165', '=B4*0.076', '=B4+C4+D4+E4+F4'],
    ['', '', '', '', '', '', ''],
    ['TOTAIS', '=SUM(B2:B4)', '=SUM(C2:C4)', '=SUM(D2:D4)', '=SUM(E2:E4)', '=SUM(F2:F4)', '=SUM(G2:G4)']
];

ForgeExcel::writeWithFormulas('impostos.xlsx', $dados);
```

---

## Dicas e Boas Práticas

### 1. Use Nomes Descritivos

```php
// BOM: Fácil de entender
['Total', '=SUM(B2:B10)']

// RUIM: Difícil de manter
['X', '=A2*B2+C2-D2/E2']
```

### 2. Documente Fórmulas Complexas

```php
$dados = [
    ['Descrição', 'Valor'],
    ['ROI (%)', '=(Receita-Custo)/Custo*100'],
    ['// Fórmula: (Receita - Custo) / Custo * 100', '']
];
```

### 3. Use Referências Absolutas Quando Necessário

```php
// Taxa de câmbio fixa
['Dólar (R$)', 5.20],
['', ''],
['Produto', 'Preço USD', 'Preço BRL'],
['Item 1', 100, '=B4*$B$1'],
['Item 2', 250, '=B5*$B$1']
```

### 4. Valide Divisões

```php
// Evita erro de divisão por zero
['Média', '=IF(B2=0,"N/A",A2/B2)']
['Média', '=IFERROR(A2/B2,"Divisão inválida")']
```

### 5. Quebre Fórmulas Complexas

```php
// BOM: Passos intermediários
$dados = [
    ['Valor', 'Desconto 10%', 'Após Desconto', 'Taxa 5%', 'Total'],
    [1000, '=A2*0.1', '=A2-B2', '=C2*0.05', '=C2+D2']
];

// RUIM: Tudo em uma fórmula
$dados = [
    ['Valor', 'Total'],
    [1000, '=((A2-(A2*0.1))+(A2-(A2*0.1))*0.05)']
];
```

---

## Limitações

### O que NÃO funciona:

❌ **Fórmulas entre abas diferentes** (precisam estar na mesma aba)  
❌ **Macros VBA** (não suportadas)  
❌ **Formatação condicional automática** (deve ser manual)  
❌ **Gráficos** (devem ser criados manualmente no Excel)  
❌ **Tabelas dinâmicas** (devem ser criadas manualmente)  

### Alternativas:

✅ Use **múltiplas abas** para organizar dados relacionados  
✅ Crie **colunas auxiliares** para cálculos intermediários  
✅ Aplique **estilos manuais** com `writeWithStyle()`  

---

## Conclusão

Com essas fórmulas, você pode criar planilhas Excel extremamente poderosas e dinâmicas! O Excel recalcula automaticamente tudo quando o arquivo é aberto.

**Próximos passos:**
- Explore o [Guia de Formatação](FORMATACAO.md)
- Veja o [Guia Completo](GUIA_COMPLETO.md)
- Execute os testes: `php test_advanced.php`

---

**Desenvolvido com ❤️ por Luan Gotardo**