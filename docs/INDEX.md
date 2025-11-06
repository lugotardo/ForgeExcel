# 📚 Documentação ForgeExcel

> **Índice completo da documentação**

---

## 🎯 Para Iniciantes

### [🚀 Quick Start](QUICKSTART.md)
**Comece aqui!** Aprenda o básico em 5 minutos.

- Instalação
- Primeiro arquivo Excel
- Casos de uso comuns
- Exemplos práticos
- Cheat sheet

---

## 📖 Documentação Completa

### [📚 Guia Completo](GUIA_COMPLETO.md)
**Referência completa** de todos os recursos do ForgeExcel.

**Conteúdo:**
- Introdução
- Instalação e configuração
- Conceitos básicos
- Operações de leitura
- Operações de escrita
- Recursos avançados
- Referência completa da API
- Exemplos práticos
- Melhores práticas
- Troubleshooting

**Ideal para:** Desenvolvedores que querem conhecer todos os recursos disponíveis.

---

## 🎨 Guias Especializados

### [🎨 Guia de Formatação](FORMATACAO.md)
**Tudo sobre estilos, cores e formatação.**

**Conteúdo:**
- Criando estilos personalizados
- Formatação de texto (negrito, itálico, sublinhado)
- Cores e fundos
- Paleta de cores predefinidas
- Alinhamento de células
- Bordas
- Aplicando estilos por linha e coluna
- Tabelas estilizadas com 5 temas
- Exemplos práticos
- Dicas e truques

**Ideal para:** Criar planilhas visualmente atraentes e profissionais.

---

### [📐 Guia de Fórmulas](FORMULAS.md)
**Tudo sobre fórmulas e cálculos automáticos.**

**Conteúdo:**
- Conceitos básicos de fórmulas
- Fórmulas matemáticas (SUM, AVERAGE, etc)
- Fórmulas estatísticas (COUNT, MAX, MIN, etc)
- Fórmulas lógicas (IF, AND, OR, etc)
- Fórmulas de texto (CONCATENATE, UPPER, etc)
- Fórmulas de data (TODAY, NOW, DATE, etc)
- Referências de células (relativas, absolutas, mistas)
- Fórmulas avançadas (VLOOKUP, SUMIFS, etc)
- Exemplos práticos completos
- Dicas e boas práticas

**Ideal para:** Criar planilhas com cálculos automáticos e dinâmicos.

---

## 📊 Recursos por Categoria

### Leitura de Arquivos
- `read()` - Leitura completa
- `readFirstSheet()` - Apenas primeira aba
- `readAllSheets()` - Todas as abas separadas
- `readInChunks()` - Processar em lotes
- `countRows()` - Contar linhas

### Escrita de Arquivos
- `write()` - Escrita simples (XLSX, CSV, ODS)
- `writeWithSheets()` - Múltiplas abas
- `writeWithStyle()` - Com formatação
- `writeWithFormulas()` - Com fórmulas
- `writeTable()` - Tabelas estilizadas
- `writeStyledSheets()` - Múltiplas abas com estilos

### Formatação e Estilos
- `createStyle()` - Criar estilo personalizado
- `colors()` - Paleta de cores
- `alignments()` - Constantes de alinhamento

### Utilitários
- `arrayToExcel()` - Converter array associativo

---

## 🎓 Níveis de Conhecimento

### Nível 1: Iniciante
📖 Leia: [Quick Start](QUICKSTART.md)
- Criar e ler arquivos Excel básicos
- Exportar dados do banco
- Importar dados para o banco

### Nível 2: Intermediário
📖 Leia: [Guia Completo](GUIA_COMPLETO.md)
- Múltiplas abas
- Processar arquivos grandes
- Arrays associativos
- Formatação básica

### Nível 3: Avançado
📖 Leia: [Guia de Formatação](FORMATACAO.md) + [Guia de Fórmulas](FORMULAS.md)
- Formatação profissional
- Fórmulas complexas
- Tabelas estilizadas
- Dashboards executivos
- Relatórios automáticos

---

## 🔍 Busca Rápida

### Preciso fazer...

**Criar um arquivo Excel simples**
→ [Quick Start - Exemplo Mais Simples](QUICKSTART.md#exemplo-mais-simples-possível)

**Exportar dados do banco de dados**
→ [Quick Start - Exportar do Banco](QUICKSTART.md#exportar-do-banco-de-dados)

**Importar dados para o banco**
→ [Quick Start - Importar para o Banco](QUICKSTART.md#importar-para-o-banco-de-dados)

**Aplicar cores e formatação**
→ [Guia de Formatação - Cores e Fundos](FORMATACAO.md#cores-e-fundos)

**Criar tabela bonita**
→ [Quick Start - Tabela Bonita](QUICKSTART.md#criar-tabela-bonita)

**Usar fórmulas do Excel**
→ [Guia de Fórmulas - Conceitos Básicos](FORMULAS.md#conceitos-básicos)

**Processar arquivo muito grande**
→ [Guia Completo - Leitura em Lotes](GUIA_COMPLETO.md#leitura-em-lotes-arquivos-grandes)

**Criar múltiplas abas**
→ [Guia Completo - Múltiplas Abas](GUIA_COMPLETO.md#múltiplas-abas)

**Fazer cálculos automáticos**
→ [Guia de Fórmulas - Fórmulas Matemáticas](FORMULAS.md#fórmulas-matemáticas)

**Criar relatório financeiro**
→ [Guia de Fórmulas - Exemplo 1](FORMULAS.md#exemplo-1-relatório-financeiro-completo)

**Aplicar estilos diferentes por linha**
→ [Guia de Formatação - Estilos por Linha](FORMATACAO.md#método-1-estilos-por-linha)

**Usar cores predefinidas**
→ [Guia de Formatação - Paleta de Cores](FORMATACAO.md#paleta-de-cores-predefinidas)

---

## 🧪 Exemplos Práticos

### Arquivos de Teste

Execute os arquivos de teste para ver exemplos funcionando:

```bash
# Testes básicos (leitura, escrita, múltiplas abas)
php test.php

# Testes avançados (formatação, fórmulas, temas)
php test_advanced.php
```

### Exemplos por Caso de Uso

**Sistema de Notas Escolares**
- [Quick Start - Sistema de Notas](QUICKSTART.md#sistema-de-notas)
- [Guia de Fórmulas - Exemplo 5](FORMULAS.md#exemplo-5-controle-de-notas-escolares)

**Dashboard Executivo**
- [Quick Start - Dashboard Executivo](QUICKSTART.md#dashboard-executivo)
- [Guia Completo - Dashboard Executivo](GUIA_COMPLETO.md#3-dashboard-executivo)

**Controle de Estoque**
- [Guia de Fórmulas - Exemplo 2](FORMULAS.md#exemplo-2-controle-de-estoque-com-alertas)

**Relatório Financeiro**
- [Guia de Fórmulas - Exemplo 1](FORMULAS.md#exemplo-1-relatório-financeiro-completo)
- [Guia de Formatação - Exemplo 1](FORMATACAO.md#exemplo-1-relatório-financeiro)

**Folha de Pagamento**
- [Guia de Fórmulas - Exemplo 3](FORMULAS.md#exemplo-3-folha-de-pagamento)

**Importação em Massa**
- [Quick Start - Importador em Massa](QUICKSTART.md#importador-em-massa)
- [Guia Completo - Exemplo 4](GUIA_COMPLETO.md#4-processar-arquivo-enorme)

---

## 📋 Referência Rápida

### Métodos Principais

| Método | Descrição | Guia |
|--------|-----------|------|
| `read()` | Ler arquivo Excel | [Guia Completo](GUIA_COMPLETO.md#read) |
| `write()` | Escrever arquivo | [Guia Completo](GUIA_COMPLETO.md#write) |
| `writeWithSheets()` | Múltiplas abas | [Guia Completo](GUIA_COMPLETO.md#writewithsheets) |
| `writeWithStyle()` | Com formatação | [Guia de Formatação](FORMATACAO.md#aplicando-estilos) |
| `writeWithFormulas()` | Com fórmulas | [Guia de Fórmulas](FORMULAS.md#conceitos-básicos) |
| `writeTable()` | Tabela estilizada | [Guia de Formatação](FORMATACAO.md#tabelas-estilizadas) |
| `readInChunks()` | Processar em lotes | [Guia Completo](GUIA_COMPLETO.md#readinchunks) |
| `createStyle()` | Criar estilo | [Guia de Formatação](FORMATACAO.md#criando-estilos) |
| `arrayToExcel()` | Converter array | [Guia Completo](GUIA_COMPLETO.md#arraytoexcel) |

### Fórmulas Suportadas

| Categoria | Fórmulas | Guia |
|-----------|----------|------|
| Matemáticas | SUM, AVERAGE, ROUND, ABS, MOD | [Guia de Fórmulas](FORMULAS.md#fórmulas-matemáticas) |
| Estatísticas | COUNT, MAX, MIN, MEDIAN, MODE | [Guia de Fórmulas](FORMULAS.md#fórmulas-estatísticas) |
| Lógicas | IF, AND, OR, NOT | [Guia de Fórmulas](FORMULAS.md#fórmulas-lógicas) |
| Texto | CONCATENATE, UPPER, LOWER, LEN | [Guia de Fórmulas](FORMULAS.md#fórmulas-de-texto) |
| Data | TODAY, NOW, DATE, YEAR, MONTH | [Guia de Fórmulas](FORMULAS.md#fórmulas-de-data) |

### Temas de Tabelas

| Tema | Cores | Preview |
|------|-------|---------|
| `blue` | Azul profissional | Header azul, linhas alternadas |
| `green` | Verde natureza | Header verde, linhas alternadas |
| `red` | Vermelho corporativo | Header vermelho, linhas alternadas |
| `orange` | Laranja vibrante | Header laranja, linhas alternadas |
| `purple` | Roxo elegante | Header roxo, linhas alternadas |

---

## ❓ Perguntas Frequentes

### Como instalar?
```bash
composer require lugotardo/forgeexel
```

### Como criar um arquivo Excel simples?
```php
ForgeExcel::write('arquivo.xlsx', $dados);
```

### Como ler com headers?
```php
$dados = ForgeExcel::read('arquivo.xlsx', true);
```

### Como aplicar formatação?
```php
ForgeExcel::writeWithStyle('arquivo.xlsx', $dados, $estilos);
```

### Como usar fórmulas?
```php
$dados = [
    ['A', 'B', 'Total'],
    [10, 20, '=A2+B2']
];
ForgeExcel::writeWithFormulas('arquivo.xlsx', $dados);
```

### Como processar arquivo grande?
```php
ForgeExcel::readInChunks('arquivo.xlsx', 1000, function($lote) {
    // Processa em lotes
});
```

---

## 🆘 Suporte

**Problemas comuns:**
- [Troubleshooting no Guia Completo](GUIA_COMPLETO.md#troubleshooting)
- [Troubleshooting Rápido no Quick Start](QUICKSTART.md#troubleshooting-rápido)

**Reportar bugs:**
- GitHub Issues

**Contato:**
- Email: luan.gotardo.dev@gmail.com

---

## 🗺️ Roadmap

### Implementado ✅
- [x] Leitura de XLSX, CSV, ODS
- [x] Escrita de XLSX, CSV, ODS
- [x] Múltiplas abas
- [x] Processamento em lotes
- [x] Formatação (cores, fontes, negrito)
- [x] Fórmulas do Excel
- [x] Tabelas estilizadas com temas
- [x] Arrays associativos
- [x] Documentação completa

### Planejado 📋
- [ ] Suporte para imagens
- [ ] Gráficos nativos
- [ ] Formatação condicional avançada
- [ ] CLI para conversões
- [ ] Validação de dados
- [ ] Proteção de células

---

## 📄 Licença

Este projeto está sob a licença MIT.

---

**Desenvolvido com ❤️ por Luan Gotardo**

💡 **Dica:** Comece pelo [Quick Start](QUICKSTART.md) e evolua gradualmente!