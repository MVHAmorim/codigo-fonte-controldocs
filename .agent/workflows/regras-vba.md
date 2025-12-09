---
description: Regras específicas para trabalhar com VBA.
---

1. Stack e Convenções de Linguagem
Qualificação Explícita: Sempre prefixe funções nativas com VBA. para evitar conflitos de referência.

Correto: VBA.Mid(), VBA.Left(), VBA.Format(), VBA.Round(), VBA.Replace().

Errado: Mid(), Left(), Format().

Tipagem Segura:

NUNCA use Val() para converter strings (falha com vírgulas). Use CDbl(), CLng() ou helpers como fnXML.ValidarValores.

Use Option Explicit obrigatoriamente.

2. Arquitetura de Dados e XML
Dicionários (Scripting.Dictionary):

Proibido: Usar Type (UDT) como item.

Obrigatório: Usar Classes DTO (ex: clsAgregacao) para armazenar dados complexos em memória.

Parsing XML:

Nunca acesse node.text diretamente (risco de Object Variable not Set). Use sempre fnXML.ValidarTag(node, xpath).

Lide com Namespaces usando fnXML.RemoverNamespaces no load ou local-name() no XPath.

3. Reuso e Helpers (Helper First Strategy)
Não reinvente a roda. Antes de codar utilitários, busque nos objetos globais existentes:

fnXML (clsFuncoesXML): Parsing, limpeza, validação de tags e conversão monetária segura.

fnSPED (clsFuncoesSPED): Geração de chaves (GerarChaveRegistro), formatação de datas/campos SPED.

Util (clsUtilitarios): Manipulação de arquivos (NomeArquivo), strings, UI (Barra de Status) e dicionários.

assImportacao (AssistenteImportacao): Reutilize métodos de negócio legados como CriarRegistro0150.

4. Estrutura de Código (SLA & Nomenclatura)
Nomenclatura: Tudo em PT-BR.

Classes/Métodos: PascalCase (CalcularImposto).

Variáveis Locais: camelCase (valorTotal).

SRP (Single Responsibility):

Métodos Public são Orquestradores (sem loops complexos, apenas delegam).

Métodos Private são Executores (loops, regras de negócio, I/O).

Tratamento de Erro: Use On Error GoTo em métodos públicos para capturar falhas e limpar a barra de status (Util.AtualizarBarraStatus False).