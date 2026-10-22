---
name: mapa-cotacoes
description: >-
  Especialista em elaboração e atualização contínua de Mapas de Cotações, equalização técnica
  e comercial, sincronização de planilhas Excel e dashboards HTML interativos para a construção
  civil. Especializado em concreto usinado e bombeamento (NBR 7212, fck, taxas e franquias) e
  aço estrutural em duas abas independentes (Corte e Dobra Industrial vs. Barra Reta 12m,
  NBR 7480 e NBR 6118), além de blocos, argamassas e insumos gerais, com simulador de
  contraproposta item a item e integração Sienge ERP.
---

# Skill: Mapa de Cotações Executivo para Construção Civil (Concreto, Aço & Insumos)

Esta skill orienta o agente a processar orçamentos técnicos e comerciais de fornecedores da construção civil (PDFs, imagens de propostas/WhatsApp, planilhas e contratos), atualizar planilhas de equalização em Excel (`.xlsx`), estruturar **Mapas de Cotação Comparativos Executivos em HTML Interativo**, auditar itens faltantes por fornecedor, equalizar condições técnicas e financeiras, e simular contrapropostas linha a linha.

---

## 1. Regra Fundamental de Identidade & Design: Proibição Absoluta de Emojis

> [!IMPORTANT]
> **NUNCA UTILIZAR EMOJIS EM HIPÓTESE ALGUMA** em mapas de cotação, dashboards, tabelas, abas, botões, gráficos, cards ou relatórios executivos.
> - **Proibido**: tratores, troféus, medalhas, lâmpadas, pranchetas, gráficos coloridos em emoji, emojis de ferramentas ou qualquer caractere emoji Unicode (`> 0x1F000`).
> - **Obrigatório**: Todo o visual deve ser estritamente executivo, sóbrio, limpo e profissional, utilizando exclusivamente:
>   - Tipografia corporativa clara com hierarquia de pesos (`Plus Jakarta Sans` / `Inter`).
>   - Badges textuais bem diagramadas (`1º Lugar`, `Líder`, `Menor Preço`, `1º Completo`, `Incompleto`, `Contrato`).
>   - Ícones vetoriais SVG minimalistas inline quando estritamente necessário (lupas, setas, ícones de impressão/cópia).

---

## 2. Escopo de Aplicação

A skill se aplica a processos de suprimentos, planejamento e compras da construtora:
1. **Concreto Usinado & Bombeamento** (Módulo Especializado): traços por $f_{ck}$, abatimento (slump), dimensão de brita, tipo de cimento, taxas de bomba, faturamento mínimo, franquias de descarga, estadia de betoneira e ensaios tecnológicos.
2. **Aço Estrutural CA-50 / CA-60 & Arame Recozido** (Módulo Dupla Modalidade): comparação segregada em **duas abas independentes** — **Aba 1: Corte e Dobra Industrial (CD)** e **Aba 2: Barra Reta (12m)** — com cálculo de custo efetivo real no canteiro (perda de pontas/sucata + mão de obra de bancada), frete (CIF vs. FOB) e arame recozido BWG 16.
3. **Argamassas e Agregados**: argamassa estabilizada (24h/36h/48h/72h), areia, brita, saibro.
4. **Alvenaria e Estrutura**: blocos cerâmicos, blocos de concreto, lajes protendidas e pré-moldados.
5. **Outros Insumos e Serviços**: esquadrias, instalações, impermeabilização, pinturas.

---

## 3. Protocolo de Execução e Atualização Contínua (7 Etapas)

```
[1. Triagem & OCR de Propostas] ➔ [2. Auditoria de Itens Faltantes] ➔ [3. Sincronização Planilha Excel]
 ➔ [4. Equalização Técnica por Disciplina] ➔ [5. Matriz de Preços Unitários]
 ➔ [6. Custos Ocultos & Equalização Comercial] ➔ [7. Dashboard HTML + Simulador de Contraproposta]
```

### Etapa 1: Triagem e Leitura de Documentos (Novas Cotações ou Atualizações)
- Ler todas as propostas na pasta da obra (PDFs, imagens `.jpeg`/`.png` enviadas por fornecedores, planilhas de equalização `.xlsx`, contratos prévios).
- Quando o usuário solicitar **"RODAR ATUALIZAÇÃO DO PACOTE"** ou adicionar novos arquivos na pasta:
  - Identificar imediatamente os arquivos novos ou revisados (`REV01`, `PT 2`, novas imagens de cotação).
  - Extrair Razão Social, CNPJ, número da cotação, contato comercial, data, validade, condição de pagamento e frete (CIF/FOB).
  - Extrair preços unitários exatos por bitola/traço/insumo, respeitando até 4 casas decimais quando informadas pela usina (ex.: `R$ 4,8041/kg`).

### Etapa 2: Auditoria de Escopo e Checklist de Itens Faltantes
- Verificar para cada fornecedor e em cada modalidade (ex.: Barra Reta vs. Corte e Dobra) quais itens da lista oficial da obra ainda estão sem preço (`null` ou não cotados).
- Sempre que solicitado ou quando houver cotações parciais, gerar uma **mensagem enxuta e direta (pronta para WhatsApp/E-mail)** listando fornecedor por fornecedor exatamente quais bitolas/itens faltam cotar (incluindo itens de referência de projeto com quantidade zero, para deixar a tabela de preços unitários 100% fechada).

### Etapa 3: Sincronização com a Planilha Excel de Equalização (`.xlsx`)
- Caso exista uma planilha Excel de equalização na pasta (ex.: `GIA-_-Equalização_Aço_-REV00.xlsx`), atualizar programaticamente via `openpyxl` as células de preços unitários dos fornecedores nas respectivas colunas (Barra Reta e Corte e Dobra), preservando todas as fórmulas (`=D*E`), formatações e abas originais.

### Etapa 4: Equalização Técnica por Disciplina

#### 4.1 Concreto Usinado (ABNT NBR 7212)
- Agrupar os traços por parâmetros equivalentes:
  - **$f_{ck}$** (Resistência característica à compressão: 15, 20, 25, 30, 35, 40, 45 MPa).
  - **Abatimento / Slump** (ex.: $10 \pm 2$ cm, $12 \pm 2$ cm, $16 \pm 2$ cm, fluido/hélice contínua $22 \pm 3$ cm ou $24 \pm 2$ cm).
  - **Dimensão Máxima do Agregado**: Brita 1, Brita 0 (pedrisco) ou mista (Brita 0+1).
  - **Tipo de Cimento / Aditivos**: CP II, CP IV, retardador de pega, plastificante, impermeabilizante.

#### 4.2 Aço Estrutural (ABNT NBR 7480 & NBR 6118) — Regra de Duas Abas Independentes
- **Obrigatoriamente separar a comparação de Aço em DUAS ABAS principais no Dashboard**:
  1. **Aba 1 — Mapa: Corte e Dobra Industrial (CD)**: compara exclusivamente as propostas de Corte e Dobra (ex.: Paranafer CD, ArcelorMittal Belgo Pronto CD, Gerdau CD, Kauefer CD, Nortefer CD).
  2. **Aba 2 — Mapa: Barra Reta (12m)**: compara exclusivamente as propostas de Barra Reta 12m / Corte Sob Medida (ex.: ArcelorMittal Barra, Gerdau Barra, Kauefer Barra, Nortefer Barra).
- **Sem Botões Artificiais de "Equalizar Arame"**:
  - **NÃO criar botão/toggle de "Equalizar Arame"** na barra de ferramentas.
  - No rodapé (`tfoot`) de cada tabela de aço, exibir de forma limpa e separada:
    - **Linha 1**: `Subtotal Aço CA-50 / CA-60 (kg)`
    - **Linha 2**: `Arame Recozido BWG 16 (kg)` (mostrando o valor cotado por cada fornecedor ou `Não cotado` quando o fornecedor não trabalhar com arame)
    - **Linha 3**: `Total Geral Cotado na Proposta`
    - **Linha 4**: `Condição de Pagamento / Frete`
    - **Linha 5**: `Classificação na Modalidade` (destacando tanto o `1º no Aço CD (S/ Arame)` quanto o `1º Completo (c/ Arame)`).
- **Auditoria de Custo Real no Canteiro (CD vs. Barra Reta)**:
  - Demonstrar matematicamente na aba de Equalização Técnica que a Barra Reta possui custos ocultos diretos no canteiro:
    - **Perda de pontas / sucata e transpasse**: `5,0% a 8,0%` (referência média: `6,5%` sobre o peso do aço).
    - **Mão de obra de bancada (corte e dobra manual na obra)**: `~R$ 0,55 a R$ 0,65/kg` (referência: `R$ 0,60/kg`).
  - Confrontar o **Custo Nominal Barra Reta + Sucata + M.O. Bancada** com o **Custo Corte e Dobra Industrial (perda zero)** para apurar a economia líquida real no canteiro.

### Etapa 5: Matriz de Preços Unitários e Regras de Seleção
- Construir tabelas comparativas dinâmicas com colunas para cada fornecedor da modalidade.
- Cada célula de preço deve conter:
  - Radio button interativo (`saas-radio`) para seleção manual item a item pelo comprador.
  - Preço unitário formatado (`R$/kg` ou `R$/m³`).
  - Badge textual de variação percentual em relação ao menor preço da linha (`Menor` em verde esmeralda ou `+X,X%` em pílula neutra/rosa).
  - Subtotal daquele item (`Volume × Preço Unitário`).
- Botões rápidos na barra de ferramentas para selecionar com 1 clique qualquer fornecedor individualmente ou o **"Menor Item a Item"** (compra mista/split de pedido).

### Etapa 6: Auditoria de Custos Ocultos e Condições Comerciais
- **No Concreto**: auditar taxas de bombeamento (bomba lança vs. estacionária, faturamento mínimo de 20 m³ a 40 m³), falso frete (carga mínima de 6 m³), estadia de betoneira, devolução ambiental e rompimento de CPs.
- **No Aço**: auditar prazos de pagamento (Antecipado vs. 15D vs. 28D/30D Boleto), frete CIF posto obra, pedido mínimo de entrega fracionada (Just-in-Time por etapa/pavimento) e certificados de qualidade de usina conforme ABNT NBR 7480.

### Etapa 7: Simulador de Contraproposta & Exportação Sienge ERP
- **Simulador de Contraproposta**:
  - Permite escolher qualquer fornecedor (tanto de Corte e Dobra quanto de Barra Reta, ou qualquer concreteira), aplicar um percentual de desconto linear (`%`) via slider, input numérico ou botões rápidos (`-2%`, `-3%`, `-5%`, `-7%`, `-10%`).
  - Exibe a **Tabela Detalhada de Preços Unitários com Desconto Simulado** linha a linha (Código Sienge, Descrição, Volume, Preço Cotado, **Preço Negociado**, Total Cotado, **Total Negociado** e Economia no Item).
  - Inclui botão **"Copiar Contraproposta"** que gera o texto formatado com todos os preços unitários alvo para envio imediato por WhatsApp/E-mail ao vendedor.
- **Integração Sienge ERP**:
  - Aba dedicada com os códigos oficiais Sienge (`01.002.0005`, `01.002.0012`, etc.), descrição, unidade, volume, fornecedor homologado na seleção ativa, preço unitário e valor total, com botão **"Copiar TSV para Sienge"** para colagem direta no ERP ou Excel.

---

## 4. Padrão Visual e Arquitetura do Dashboard HTML Interativo

O mapa de cotações deve ser entregue como um arquivo HTML único, autocontido (CSS + JS + Logo Base64 embutidos), moderno, responsivo e 100% interativo:

### 4.1 Identidade Visual Corporativa
- **Paleta de Cores**:
  - Primária: `#0a2540` / `#163970` (Azul Marinho Corporativo).
  - Destaques: `#0891b2` (Ciano/Turquesa para Corte e Dobra), `#2563eb` (Azul Real para Barra Reta / Concreto) e `#0284c7` (Azul Celeste para Bombeamento).
  - Fundo geral: `#f8fafc` (`bg-slate-50`).
  - Superfícies: Cartões brancos `rounded-2xl`, bordas `border-slate-200/80` e sombras sutis `shadow-xs`.
  - Semáforo: Verde `#059669` (Menor preço/atendido), Âmbar `#d97706` (Atenção/escopo parcial/contrato), Rosa `#e11d48` (Incompleto/custo oculto elevado).
- **Tipografia e Precisão Numérica**:
  - Fonte: `Plus Jakarta Sans` via Google Fonts.
  - Classe obrigatória `num-tabular` (`font-feature-settings: "tnum" 1; font-variant-numeric: tabular-nums;`) em todos os valores monetários, percentuais e quantitativos.
  - Volumes e quantidades exibidos sempre com 2 casas decimais (`18.999,00 kg`, `12.672,00 kg`, `2.942,00 m³`).

### 4.2 Estrutura Completa de Abas

#### Para Mapas de Aço Estrutural (8 Abas):
1. **Mapa: Corte e Dobra (CD)**: Tabela comparativa exclusiva dos fornecedores de Corte e Dobra Industrial + 3 Cards inferiores (Gráfico Donut de distribuição do custo CD, Composição de Bitolas & Inspetor Técnico de Linha, e Resumo do Cenário CD).
2. **Mapa: Barra Reta (12m)**: Tabela comparativa exclusiva dos fornecedores de Barra Reta 12m + 3 Cards inferiores dedicados (Gráfico Donut Barra, Composição de Bitolas & Inspetor Técnico Barra, e Resumo do Cenário Barra com alerta de custo real de canteiro).
3. **Resumo Executivo & Ranking Geral**: Tabela consolidada com todas as propostas das duas modalidades ordenadas e comparadas tanto pelo **Valor Cotado Nominal** quanto pelo **Custo Real Estimado no Canteiro**.
4. **Simulações & Negociação**: Simulador interativo com tabela completa de bitolas e preços unitários descontados + cópia de contraproposta.
5. **Equalização Técnica (CD vs. Barra)**: Matriz comparativa de engenharia (perda de sucata 6,5%, mão de obra de armadores no canteiro, espaço físico, produtividade de formas, segurança NR-18 e balanço financeiro líquido).
6. **Análise Técnica & Normas**: Auditoria de conformidade do projeto estrutural vs. oferta comercial + cards de Massa Linear Teórica e Tolerâncias da **ABNT NBR 7480** (`0,154 kg/m` para `5,0mm` até `3,853 kg/m` para `25,0mm`).
7. **Importação Sienge ERP**: Exportação TSV da seleção ativa (Corte e Dobra ou Barra Reta).
8. **Gráficos Comparativos**: Gráficos horizontais de barras comparando todas as propostas e a distribuição física do peso por bitola.

#### Para Mapas de Concreto Usinado (7 Abas):
1. **Mapa Comparativo**: Tabela dinâmica de traços com controle de Bombeamento e Ajuste a Valor Presente (AVP) + 3 Cards inferiores.
2. **Resumo Executivo & Ranking**: Classificação comercial e histórico de contratos.
3. **Simulações & Negociação**: Simulador de desconto com planilha detalhada traço a traço.
4. **Equalização de Bombeamento**: Comparativo de bombas, franquias e taxas ocultas.
5. **Análise Técnica & Riscos**: Matriz de $f_{ck}$, slump, brita e controle tecnológico NBR 7212.
6. **Importação Sienge ERP**: Tabela padrão Sienge pronta para cópia.
7. **Gráficos Comparativos**: Custo total por concreteira e distribuição de volume por elemento estrutural.

### 4.3 Impressão A4 Paisagem & Exportações
- `@page { size: A4 landscape; margin: 6mm; }` e `print-color-adjust: exact !important;` para impressão executiva em PDF fiel à tela.
- `tfoot { display: table-row-group !important; }` para evitar repetição do rodapé de totais em páginas intermediárias.
- Exportação funcional para CSV/Excel e botão nativo `Imprimir / PDF`.

---

## 5. Checklist de Verificação Antes da Entrega

Sempre executar uma verificação automatizada (via Python) no HTML gerado antes de concluir:
1. **Zero Emojis**: `len([c for c in html if ord(c) > 0x1F000]) == 0`.
2. **Integridade de Sintaxe JavaScript**: verificar balanceamento exato de chaves `{ }`, parênteses `( )` e colchetes `[ ]` dentro da tag `<script>` para garantir que todas as tabelas e gráficos sejam renderizados no `DOMContentLoaded`.
3. **Consistência Matemática**: validar se a soma de `Quantidade × Preço Unitário` de cada item bate exatamente com o Subtotal, Total Geral com Arame/Bomba e Planilha Excel (`.xlsx`).
