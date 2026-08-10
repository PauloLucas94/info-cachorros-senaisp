# BI — Cachorros nas Unidades SENAI-SP

Guia de construção do dashboard a partir da aba **BI_Base** da planilha `Informações Cachorros SENAI.xlsx`.

**Grão da tabela:** 1 linha = 1 cachorro. 16 linhas, 35 colunas, 11 unidades (CFP).
As abas originais (`Resumo`, `Base Consolidada`, `Resumo por CFP`, `Pontos de Atenção`) não foram alteradas.

> **Versão 4** — ajustada após a remoção das colunas `Unidade`, `Cidade`, `UF`, `Regiao Administrativa` e `Tem Foto`; a inclusão de `Vacinas Recebidas`, `URL Foto`, `Qual Documento`, `URL Carteira` e `Responsavel Completo`; e a **conversão dos 8 campos de conformidade para Verdadeiro/Falso** no Power Query. Todas as medidas abaixo já estão na sintaxe booleana.
>
> Após varredura campo a campo das 4 abas originais, **os 23 campos da `Base Consolidada` estão integralmente representados na BI_Base**. As abas `Pontos de Atenção`, `Resumo por CFP` e `Resumo` são 100% deriváveis — nenhuma informação exclusiva ficou de fora.

---

## 1. Importação no Power BI

1. **Obter dados → Pasta de trabalho do Excel** → selecione o arquivo.
2. No Navegador, marque **apenas `tbBI_Base`** (é uma Tabela nomeada — importa mais estável que a planilha inteira).
3. **Transformar Dados** e confira os tipos:

| Coluna | Tipo | Categoria de dados |
|---|---|---|
| `CFP` | Texto | — |
| `Latitude` | Número decimal | Latitude |
| `Longitude` | Número decimal | Longitude |
| `Idade (anos)` | Número decimal | — |
| `URL Foto` | Texto | URL da Web |
| `URL Carteira` | Texto | URL da Web |
| `Vacinas Recebidas` | Texto | — |
| `Qual Documento` | Texto | — |
| `Responsavel Completo` | Texto | — |
| `Castrado`, `Vacinado`, `Microchipado`, `Abrigo Adequado`, `Documento SENAI`, `Possui Doenca`, `Medicamento Continuo`, `Tem Carteira Vacinacao` | **Verdadeiro/Falso** | — |
| `ID`, `Criterios Atendidos`, `Total Criterios`, `Qtd Pendencias` | Número inteiro | — |
| `Indice Conformidade` | Número decimal | — |

> **Sobre os campos de conformidade.** Na planilha eles estão gravados como 0/1, mas são convertidos para **Verdadeiro/Falso** no Power Query. É mais correto semanticamente e habilita ícones na formatação condicional. Em troca, **`SUM` deixa de funcionar sobre eles** — DAX não soma booleano. Por isso as contagens da seção 3.2 usam `CALCULATE` + `COUNTROWS` em vez de `SUM`.
>
> As colunas `Criterios Atendidos`, `Total Criterios` e `Qtd Pendencias` continuam **inteiras** — é o que mantém o `Índice Conformidade Geral` e as medidas de pendência funcionando com `SUM`.
>
> Se em algum momento você reverter os 8 campos para inteiro, troque cada `CALCULATE ( COUNTROWS ( ... ), col = TRUE )` por `SUM ( col )`.

4. Em **Modelagem**, marque `Latitude` e `Longitude` com **Resumir por: Não resumir**.

### Ordenação correta de `Faixa Etaria`

Crie uma coluna calculada e use *Classificar por coluna*:

```dax
Ordem Faixa Etaria =
SWITCH (
    tbBI_Base[Faixa Etaria],
    "Filhote (< 1 ano)", 1,
    "Jovem (1-2 anos)", 2,
    "Adulto (3-7 anos)", 3,
    "Idoso (8+ anos)", 4,
    99
)
```

Idem para `Classificacao`:

```dax
Ordem Classificacao =
SWITCH ( tbBI_Base[Classificacao], "Crítico", 1, "Atenção", 2, "Conforme", 3, 99 )
```

---

## 2. Opcional — devolver o nome da unidade sem sujar a tabela de fatos

Com `Unidade` removida, todo rótulo de gráfico mostrará o código (`5.12`, `9.14`). Se quiser os nomes de volta, **não recoloque a coluna na BI_Base** — crie uma tabela de dimensão dentro do próprio Power BI, que é o padrão correto de modelagem.

**Página Inicial → Inserir dados**, nomeie como `dUnidade` e cole:

| CFP | Unidade | Cidade | Regiao |
|---|---|---|---|
| 1.02 | Vila Alpina | São Paulo | Metropolitana |
| 1.28 | Guarulhos | Guarulhos | Metropolitana |
| 1.36 | Barueri | Barueri | Metropolitana |
| 2.60 | Registro | Registro | Registro |
| 3.01 | Taubaté | Taubaté | Vale do Paraíba |
| 5.08 | Itatiba | Itatiba | Campinas |
| 5.12 | Sumaré | Sumaré | Campinas |
| 5.69 | Paulínia | Paulínia | Campinas |
| 5.94 | Iracemápolis | Iracemápolis | Campinas |
| 7.92 | Lençóis Paulista | Lençóis Paulista | Bauru |
| 9.14 | Pres. Prudente | Presidente Prudente | Pres. Prudente |

Depois relacione `dUnidade[CFP]` **1 → N** com `tbBI_Base[CFP]`. Passe a usar `dUnidade[Unidade]` nos eixos e segmentações.

**Se optar por não criar a `dUnidade`**, tudo abaixo funciona igual — é só usar `tbBI_Base[CFP]` onde o texto menciona unidade. As medidas da seção 2.6 já vêm nas duas versões.

---

## 3. Medidas DAX

Crie uma tabela vazia chamada `_Medidas` (Inserir dados → tabela em branco) e coloque tudo nela para organizar o painel de campos.

### 3.1 Contagens base

```dax
Total Cachorros = COUNTROWS ( tbBI_Base )

Total Unidades = DISTINCTCOUNT ( tbBI_Base[CFP] )

Média Cachorros por Unidade =
DIVIDE ( [Total Cachorros], [Total Unidades] )

Idade Média = AVERAGE ( tbBI_Base[Idade (anos)] )
```

### 3.2 Conformidade — contagens

```dax
Vacinados     = CALCULATE ( COUNTROWS ( tbBI_Base ), tbBI_Base[Vacinado] = TRUE )
Castrados     = CALCULATE ( COUNTROWS ( tbBI_Base ), tbBI_Base[Castrado] = TRUE )
Microchipados = CALCULATE ( COUNTROWS ( tbBI_Base ), tbBI_Base[Microchipado] = TRUE )
Com Abrigo    = CALCULATE ( COUNTROWS ( tbBI_Base ), tbBI_Base[Abrigo Adequado] = TRUE )
Com Documento = CALCULATE ( COUNTROWS ( tbBI_Base ), tbBI_Base[Documento SENAI] = TRUE )

Não Vacinados     = [Total Cachorros] - [Vacinados]
Não Castrados     = [Total Cachorros] - [Castrados]
Não Microchipados = [Total Cachorros] - [Microchipados]
Sem Abrigo        = [Total Cachorros] - [Com Abrigo]
Sem Documento     = [Total Cachorros] - [Com Documento]
```

### 3.3 Conformidade — percentuais

```dax
% Vacinados     = DIVIDE ( [Vacinados],     [Total Cachorros] )
% Castrados     = DIVIDE ( [Castrados],     [Total Cachorros] )
% Microchipados = DIVIDE ( [Microchipados], [Total Cachorros] )
% Com Abrigo    = DIVIDE ( [Com Abrigo],    [Total Cachorros] )
% Com Documento = DIVIDE ( [Com Documento], [Total Cachorros] )
```

Formate todas como **Porcentagem, 0 ou 1 casa**.

### 3.4 Índice geral de conformidade

Média dos 5 critérios (castrado, vacinado, microchipado, abrigo, documento). Funciona em qualquer contexto de filtro:

```dax
Índice Conformidade Geral =
DIVIDE (
    SUM ( tbBI_Base[Criterios Atendidos] ),
    SUM ( tbBI_Base[Total Criterios] )
)
```

### 3.5 Pendências e criticidade

```dax
Total Pendências = SUM ( tbBI_Base[Qtd Pendencias] )

Cachorros com Pendência =
CALCULATE ( [Total Cachorros], tbBI_Base[Qtd Pendencias] > 0 )

% Cachorros com Pendência =
DIVIDE ( [Cachorros com Pendência], [Total Cachorros] )

Cachorros Críticos =
CALCULATE ( [Total Cachorros], tbBI_Base[Classificacao] = "Crítico" )

Unidades com Pendência =
CALCULATE ( [Total Unidades], tbBI_Base[Qtd Pendencias] > 0 )
```

### 3.6 Ranking e destaques

**Versão A — sem `dUnidade`** (rótulo sai como código CFP):

```dax
Unidade Mais Crítica =
VAR t =
    ADDCOLUMNS (
        VALUES ( tbBI_Base[CFP] ),
        "@pend", CALCULATE ( SUM ( tbBI_Base[Qtd Pendencias] ) )
    )
RETURN
    CONCATENATEX (
        TOPN ( 1, t, [@pend], DESC, tbBI_Base[CFP], ASC ),
        "CFP " & tbBI_Base[CFP] & " (" & [@pend] & " pendências)",
        ", "
    )

Ranking Unidade por Conformidade =
IF (
    HASONEVALUE ( tbBI_Base[CFP] ),
    RANKX ( ALLSELECTED ( tbBI_Base[CFP] ), [Índice Conformidade Geral],, DESC, DENSE )
)
```

**Versão B — com `dUnidade`** (rótulo com o nome da unidade):

```dax
Unidade Mais Crítica =
VAR t =
    ADDCOLUMNS (
        VALUES ( dUnidade[Unidade] ),
        "@pend", CALCULATE ( SUM ( tbBI_Base[Qtd Pendencias] ) )
    )
RETURN
    CONCATENATEX (
        TOPN ( 1, t, [@pend], DESC, dUnidade[Unidade], ASC ),
        dUnidade[Unidade] & " (" & [@pend] & " pendências)",
        ", "
    )

Ranking Unidade por Conformidade =
IF (
    HASONEVALUE ( dUnidade[Unidade] ),
    RANKX ( ALLSELECTED ( dUnidade[Unidade] ), [Índice Conformidade Geral],, DESC, DENSE )
)
```

### 3.7 Documentação e anexos

```dax
Com Carteira Anexada =
CALCULATE ( COUNTROWS ( tbBI_Base ), tbBI_Base[Tem Carteira Vacinacao] = TRUE )

% Com Carteira Anexada =
DIVIDE ( [Com Carteira Anexada], [Total Cachorros] )

Vacinados sem Carteira Anexada =
CALCULATE (
    COUNTROWS ( tbBI_Base ),
    tbBI_Base[Vacinado] = TRUE,
    tbBI_Base[Tem Carteira Vacinacao] = FALSE
)
```

> A última mede o descompasso entre "declarou vacinado" e "anexou comprovante" — hoje são 3 casos (Butti, Roberto Carlos, Jéssica).

### 3.8 Antirrábica declarada

Como `Vacinas Recebidas` é texto livre, dá para medir menção à antirrábica sem criar coluna nova:

```dax
Menciona Antirrábica =
CALCULATE (
    [Total Cachorros],
    FILTER (
        tbBI_Base,
        CONTAINSSTRING ( tbBI_Base[Vacinas Recebidas], "rábica" )
            || CONTAINSSTRING ( tbBI_Base[Vacinas Recebidas], "Raiva" )
    )
)

Vacinado sem Antirrábica Declarada =
CALCULATE (
    [Total Cachorros],
    FILTER (
        tbBI_Base,
        tbBI_Base[Vacinado] = TRUE
            && NOT (
                CONTAINSSTRING ( tbBI_Base[Vacinas Recebidas], "rábica" )
                    || CONTAINSSTRING ( tbBI_Base[Vacinas Recebidas], "Raiva" )
            )
    )
)
```

> `CONTAINSSTRING` não diferencia maiúsculas de minúsculas, então cobre "Antirrábica", "antirrábica" e "anti-rábica". Hoje a segunda medida retorna 4: Clarinha e Leide (5.94), que listam a polivalente sem citar raiva; José Ephim Mindlin (1.36), que respondeu só "Todas"; e Princesa (5.08), que remete à carteira anexada. **É indício de resposta incompleta no Forms, não prova de ausência de vacina.** Use como lista de conferência, nunca como indicador final.

### 3.9 Formatação condicional por cor

Use em *Formatação condicional → Cor do plano de fundo → Formato por: Campo*:

```dax
Cor Conformidade =
VAR i = [Índice Conformidade Geral]
RETURN
    SWITCH (
        TRUE (),
        i >= 0.8,  "#2E7D32",   -- verde
        i >= 0.5,  "#F9A825",   -- âmbar
        "#C62828"                -- vermelho
    )
```

### 3.10 Texto dinâmico para o cabeçalho

```dax
Título Dinâmico =
"Cachorros nas Unidades SENAI-SP — "
    & [Total Cachorros] & " animais em "
    & [Total Unidades] & " unidades | Conformidade: "
    & FORMAT ( [Índice Conformidade Geral], "0%" )
```

---

## 4. Roteiro do dashboard

**Paleta sugerida** (institucional SENAI): azul `#1F3864` (primário), vermelho `#C8102E` (alerta), cinza `#F2F2F2` (fundo), verde `#2E7D32` (conforme), âmbar `#F9A825` (atenção).

> Onde estiver escrito **`[Unidade]`**, use `dUnidade[Unidade]` se tiver criado a dimensão da seção 2, ou `tbBI_Base[CFP]` se não tiver.

### Página 1 — Visão Geral

**Faixa de KPIs (topo, 6 cartões):**

| Cartão | Medida |
|---|---|
| Total de cachorros | `Total Cachorros` |
| Unidades com cachorros | `Total Unidades` |
| % Vacinados | `% Vacinados` |
| % Castrados | `% Castrados` |
| % Microchipados | `% Microchipados` |
| Índice de conformidade | `Índice Conformidade Geral` |

**Corpo:**

- **Mapa** (esquerda, ~40% da largura) — `Latitude`/`Longitude`, tamanho da bolha = `Total Cachorros`, cor = `Cor Conformidade`. Dica de ferramenta: `[Unidade]`, `Nome do Cachorro`, `Pendencias`.
- **Barras horizontais** (direita) — Eixo Y `[Unidade]`, valor `Total Cachorros`, ordenado decrescente.
- **Gráfico de rosca** — `Classificacao` (Conforme / Atenção / Crítico).
- **Barras empilhadas 100%** — Eixo `[Unidade]`, valores `Vacinados` vs `Não Vacinados`.

**Segmentações (painel lateral retrátil):** `Porte`, `Sexo`, `Faixa Etaria`, `Classificacao`, `Vinculo com a Unidade`. Acrescente `dUnidade[Regiao]` se criou a dimensão.

### Página 2 — Conformidade e Pendências

- **Matriz** — Linhas: `[Unidade]` › `Nome do Cachorro`. Valores: as 5 medidas de contagem (`Castrados`, `Vacinados`, `Microchipados`, `Com Abrigo`, `Com Documento`). Com os campos booleanos **não arraste as colunas direto para Valores** — o Power BI só oferece *Contagem* e o resultado fica 1 em toda linha. Use as medidas, e aplique formatação condicional por ícones (✔ verde para 1, ✖ vermelho para vazio/0).
- **Gráfico de barras** — `Qtd Pendencias` por `[Unidade]`, cor vermelha, ordenado decrescente.
- **Cartão de múltiplas linhas** — `Cachorros com Pendência`, `Unidades com Pendência`, `Total Pendências`, `Cachorros Críticos`.
- **Tabela de detalhe** — `Nome do Cachorro`, `[Unidade]`, `Pendencias`, `Indice Conformidade`. Filtrada em `Qtd Pendencias > 0`.
- **Funil de critérios** — categoria = nome do critério, valor = quantos atendem. Mostra visualmente que documento SENAI é o gargalo (3/16).
- **Tabela de documentos** — `[Unidade]`, `Nome do Cachorro`, `Documento SENAI`, `Qual Documento`. Deixa explícito que os 3 positivos não são equivalentes: 1.02 tem "Crachá e RGA", enquanto 1.28 tem apenas um e-mail de analista. Use ao apresentar o indicador de formalização.

### Página 3 — Saúde e Vacinação

- **Cartões** — `% Vacinados`, `% Com Carteira Anexada`, `Vacinados sem Carteira Anexada`, `Vacinado sem Antirrábica Declarada`.
- **Tabela** — `Nome do Cachorro`, `[Unidade]`, `Vacinado`, `Vacinas Recebidas`, `Tem Carteira Vacinacao`, `URL Carteira`. É a página onde o texto livre das vacinas ganha utilidade, com o comprovante clicável ao lado. Só 10 dos 16 têm carteira anexada; nos outros a coluna mostra "Sem carteira anexada".
- **Barras** — `Acompanhamento Veterinario` (Semestral / Quando necessário).
- **Cartão de destaque** — filtro `Possui Doenca` é **Verdadeiro**, exibindo `Qual Doenca` (hoje só Roberto Carlos, 7.92).

### Página 4 — Perfil dos Animais

- **Colunas agrupadas** — `Faixa Etaria` × `Porte`.
- **Rosca** — `Sexo`.
- **Barras** — `Permanencia` (Integral 24h / Circula livremente / Apenas durante o dia).
- **Cartão** — `Idade Média`.

> `Raca Normalizada` ficou com 15 SRD contra 1 Doberman — não vale um visual próprio, use no máximo como coluna de tabela.

### Página 5 — Detalhe por Unidade (drill-through)

Configure **drill-through** por `[Unidade]`. A página traz:

- Cabeçalho com `Título Dinâmico`.
- Cartões: total de cachorros da unidade, índice de conformidade, `Ranking Unidade por Conformidade`.
- Tabela com nome, sexo, porte, idade, permanência, pendências.
- **Responsável:** use `Responsavel (curto)` no eixo de gráficos e agrupamentos, e `Responsavel Completo` na tabela de detalhe — a versão longa traz cargo e equipe de apoio (o registro de 1.36 tem 344 caracteres descrevendo os 4 colaboradores e seus turnos).
- Colunas `URL Foto` e `URL Carteira` como links clicáveis para o SharePoint.

---

## 5. O que os dados já mostram

| Indicador | Valor | Leitura |
|---|---|---|
| Cachorros / unidades | 16 em 11 CFPs | Média de 1,45 por unidade |
| Vacinados | 13 (81%) | 3 pendentes: Nina, Nino (5.12) e Negão (5.69) |
| Castrados | 13 (81%) | 3 pendentes: Butti (2.60), Nino (5.12), Negão (5.69) |
| Com abrigo adequado | 13 (81%) | 3 sem abrigo: Butti, Nina, Nino |
| **Microchipados** | **5 (31%)** | Maior gargalo operacional |
| **Documento vinculado ao SENAI** | **3 (19%)** | Gargalo mais crítico — só 1.02 e 1.28 possuem |
| Classificação | 5 conformes, 7 em atenção, 4 críticos | — |

**Unidades críticas:** 5.12 Sumaré (Nina e Nino acumulam o maior número de pendências) e 5.69 Paulínia (Negão, animal comunitário, sem vacina, castração ou microchip).

**Narrativa sugerida:** o cuidado sanitário básico (vacina, castração, abrigo) está em ~81%, mas a **formalização** — microchipagem e documento vinculado ao SENAI — está muito atrás. Esse é o eixo principal do painel.

---

## 6. Ressalvas antes de publicar

1. **As coordenadas não vieram da fonte.** Não havia lat/long na planilha original; são centroides aproximados de município, não o endereço da unidade. Valide contra o cadastro oficial antes de apresentar o mapa.
2. **Princesa (5.08) está como vacinada, mas há vacina pendente.** O texto informa que a vacina importada foi encomendada e ainda será aplicada.
3. **`Documento SENAI` verdadeiro soma coisas diferentes.** 1.02 tem crachá e RGA; 1.28 tem um e-mail de analista. A coluna `Qual Documento` agora expõe isso — mostre-a sempre que apresentar o indicador de formalização.
4. **`Tem Carteira Vacinacao` mede anexo, não existência.** Célula vazia virou Falso. São 10 anexos em 16 cachorros.

   **Nota sobre segmentações booleanas:** dependendo da versão e do idioma do Power BI Desktop, o filtro exibe "True/False" em vez de "Verdadeiro/Falso". Se for expor esses campos como segmentação num painel institucional, considere criar colunas texto "Sim/Não" só para essa finalidade — as medidas continuam usando os booleanos.
5. **O índice de conformidade é convenção interna** — 5 critérios de peso igual, cortes em 80%/50%. Não é norma SENAI.
6. **Juca e Branca (9.14) apontam para a mesma foto.** Provavelmente não existe foto do Juca.
7. **Com n=16, cada cachorro vale 6,25 pontos percentuais.** Mostre sempre a contagem absoluta ao lado do percentual.
8. **Não há coluna de data na base** — não é possível montar dimensão de tempo nem medir evolução.

---

## 7. Aba BI_Inconsistencias

Documenta 11 divergências entre as abas. Duas merecem atenção:

1. **CFP 5.12 na aba `Resumo por CFP`** — a linha tem 8 valores contra 7 colunas das demais, deslocando os flags. A `Base Consolidada` é a fonte correta.
2. **CFP 5.69 (`Animal comunitário`)** — texto no lugar de SIM/NÃO. Em `BI_Base` virou a coluna categórica `Vinculo com a Unidade`.

As contagens da aba `Resumo` (16 / 13 / 13 / 5) conferem 100% com `BI_Base` — use como validação do modelo.
