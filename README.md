# 📂 CSharp File Converters

Um conjunto de utilitários em **C#** para conversão de arquivos entre diferentes formatos:

- **XLSX → CSV**
- **CSV → XLSX**
- **MDB → CSV** (exporta cada tabela para um arquivo separado)
- **CSV → XML**
- **XML → CSV**

---

## 📋 Pré-requisitos

Antes de compilar, instale os pacotes **NuGet** necessários:

```bash
dotnet add package ClosedXML
dotnet add package CsvHelper
```
Para usar MDB → CSV, é necessário:
- Windows
- Microsoft Access Database Engine (ACE ou Jet) instalado

## 🚀 Utilização
XLSX → CSV
```bash dotnet run -- xlsx2csv arquivo.xlsx arquivo.csv ```

CSV → XLSX
```bash dotnet run -- csv2xlsx arquivo.csv arquivo.xlsx ```

MDB → CSV (exporta todas as tabelas)
```bash dotnet run -- mdb2csv arquivo.mdb pasta_saida ```

CSV → XML
```bash dotnet run -- csv2xml arquivo.csv arquivo.xml ```

XML → CSV
```bash dotnet run -- xml2csv arquivo.xml arquivo.csv ```

## 📌 Observações
- O comando mdb2csv cria um arquivo CSV para cada tabela encontrada no banco .mdb.

- O xml2csv espera um formato simples, como:
```bash

Copiar
Editar
<Rows>
  <Row>
    <ColA>Valor 1</ColA>
    <ColB>Valor 2</ColB>
  </Row>
</Rows>
```
Todos os nomes de arquivos de saída que contêm caracteres inválidos serão automaticamente sanitizados.
