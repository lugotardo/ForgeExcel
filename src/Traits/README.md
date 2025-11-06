# 📁 Traits - Organização do ForgeExcel

> **Estrutura modular para facilitar manutenção e extensão**

---

## 📖 Visão Geral

Para manter o código organizado e fácil de manter, o ForgeExcel foi dividido em **5 traits** especializados, cada um responsável por um conjunto específico de funcionalidades.

A classe principal `ForgeExcel.php` agora tem apenas **36 linhas**, importando os traits necessários.

---

## 🗂️ Estrutura dos Traits

### 1️⃣ ReadTrait.php (313 linhas)
**Responsabilidade:** Leitura de arquivos Excel

**Métodos:**
- `read()` - Lê arquivo completo
- `readFirstSheet()` - Lê apenas primeira aba
- `readAllSheets()` - Lê todas as abas separadamente
- `countRows()` - Conta linhas do arquivo
- `readInChunks()` - Processa arquivo em lotes (chunks)

**Exemplo:**
```php
$dados = ForgeExcel::read('arquivo.xlsx', true);
```

---

### 2️⃣ WriteTrait.php (145 linhas)
**Responsabilidade:** Escrita básica de arquivos Excel

**Métodos:**
- `write()` - Escreve arquivo Excel/CSV/ODS
- `writeWithSheets()` - Cria arquivo com múltiplas abas

**Exemplo:**
```php
ForgeExcel::write('saida.xlsx', $dados);
```

---

### 3️⃣ StyleTrait.php (396 linhas)
**Responsabilidade:** Formatação e estilos

**Métodos:**
- `createStyle()` - Cria estilo personalizado
- `writeWithStyle()` - Escreve com formatação
- `writeTable()` - Cria tabelas com temas
- `writeStyledSheets()` - Múltiplas abas estilizadas
- `colors()` - Paleta de cores predefinidas
- `alignments()` - Constantes de alinhamento

**Exemplo:**
```php
$estilo = ForgeExcel::createStyle(['bold' => true, 'color' => 'FF0000']);
ForgeExcel::writeTable('tabela.xlsx', $dados, 'blue');
```

---

### 4️⃣ FormulaTrait.php (76 linhas)
**Responsabilidade:** Fórmulas do Excel

**Métodos:**
- `writeWithFormulas()` - Escreve arquivo com fórmulas Excel

**Exemplo:**
```php
$dados = [
    ['A', 'B', 'Total'],
    [10, 20, '=A2+B2']
];
ForgeExcel::writeWithFormulas('calculos.xlsx', $dados);
```

---

### 5️⃣ UtilityTrait.php (52 linhas)
**Responsabilidade:** Métodos utilitários

**Métodos:**
- `arrayToExcel()` - Converte array associativo para Excel

**Exemplo:**
```php
$usuarios = [['nome' => 'João', 'email' => 'joao@email.com']];
$excel = ForgeExcel::arrayToExcel($usuarios);
```

---

## 🎯 Vantagens da Separação

### ✅ Manutenibilidade
Cada arquivo tem uma responsabilidade clara e específica.

### ✅ Legibilidade
Fácil encontrar e entender onde cada funcionalidade está implementada.

### ✅ Extensibilidade
Adicionar novos recursos é simples - basta criar um novo trait ou estender um existente.

### ✅ Testabilidade
Cada trait pode ser testado independentemente.

### ✅ Tamanho Gerenciável
Nenhum arquivo tem mais de 400 linhas.

---

## 📊 Comparação

### Antes da Refatoração
```
src/ForgeExcel.php: 925 linhas
```

### Depois da Refatoração
```
src/ForgeExcel.php:              36 linhas (classe principal)
src/Traits/ReadTrait.php:       313 linhas (leitura)
src/Traits/WriteTrait.php:      145 linhas (escrita)
src/Traits/StyleTrait.php:      396 linhas (formatação)
src/Traits/FormulaTrait.php:     76 linhas (fórmulas)
src/Traits/UtilityTrait.php:     52 linhas (utilitários)
────────────────────────────────────────────────
Total:                         1018 linhas
```

---

## 🔧 Como Adicionar Novos Recursos

### 1. Identificar a categoria
Determine qual trait é mais apropriado para o novo recurso.

### 2. Adicionar ao trait existente
```php
// Em StyleTrait.php, por exemplo
public static function newStyleMethod(): void
{
    // Implementação
}
```

### 3. Ou criar novo trait (se necessário)
```php
// src/Traits/ChartTrait.php
namespace Lugotardo\Forgeexel\Traits;

trait ChartTrait
{
    public static function createChart(): void
    {
        // Implementação
    }
}
```

### 4. Importar na classe principal
```php
// Em ForgeExcel.php
use Lugotardo\Forgeexel\Traits\ChartTrait;

class ForgeExcel
{
    use ReadTrait;
    use WriteTrait;
    use StyleTrait;
    use FormulaTrait;
    use UtilityTrait;
    use ChartTrait; // Novo trait
}
```

---

## 📚 Boas Práticas

### ✅ Um trait = Uma responsabilidade
Cada trait deve ter um propósito claro e específico.

### ✅ Métodos estáticos
Todos os métodos públicos devem ser estáticos para manter a interface consistente.

### ✅ Documentação completa
Cada método deve ter PHPDoc completo em português.

### ✅ Exemplos de uso
Inclua exemplos práticos na documentação de cada método.

### ✅ Tratamento de erros
Use Exception com mensagens claras e descritivas.

---

## 🧪 Testando os Traits

Execute os testes para garantir que tudo funciona:

```bash
# Testes básicos
php test.php

# Testes avançados
php test_advanced.php
```

---

## 🤝 Contribuindo

Ao contribuir com novos recursos:

1. Identifique o trait apropriado
2. Mantenha o padrão de código existente
3. Adicione documentação completa
4. Crie testes para o novo recurso
5. Atualize este README se necessário

---

## 📖 Referências

- **[Guia Completo](../../docs/GUIA_COMPLETO.md)** - Documentação completa
- **[Guia de Formatação](../../docs/FORMATACAO.md)** - Estilos e cores
- **[Guia de Fórmulas](../../docs/FORMULAS.md)** - Fórmulas Excel

---

**Desenvolvido com ❤️ por Luan Gotardo**