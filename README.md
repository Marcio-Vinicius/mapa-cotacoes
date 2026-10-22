# Skill: Mapa de Cotações Executivo para Construção Civil (`mapa-cotacoes`)

Skill avançada para **Antigravity / Agentes de IA** voltada à elaboração e atualização contínua de **Mapas de Cotações Comparativos Executivos** para a construção civil (Concreto Usinado, Bombeamento, Aço Estrutural em Barra Reta e Corte & Dobra, Argamassas, Blocos e Insumos Gerais).

---

## Principais Recursos da Skill

1. **Zero Emojis (Padrão Estritamente Executivo)**:
   - Interface SaaS corporativa limpa (`Plus Jakarta Sans`, números tabulares `num-tabular`, badges textuais e ícones SVG vetoriais minimalistas).
2. **Módulo Especializado de Aço Estrutural (ABNT NBR 7480 / NBR 6118)**:
   - **Duas Abas Principais Independentes**:
     - **Aba 1 — Corte e Dobra Industrial (CD)**: comparação exclusiva entre fornecedores de corte e dobra + subtotal de aço + arame recozido BWG 16.
     - **Aba 2 — Barra Reta (12m)**: comparação exclusiva entre fornecedores de barra reta / corte sob medida.
   - **Equalização Técnica Real de Canteiro**: cálculo automático de perda por sucata/pontas (`6,5%`) + mão de obra de bancada (`R$/kg`), demonstrando a diferença real entre o valor nominal da barra reta e o custo efetivo no canteiro.
3. **Módulo Especializado de Concreto Usinado & Bombeamento (ABNT NBR 7212)**:
   - Equalização por $f_{ck}$, slump, dimensão de brita e aditivos.
   - Discriminação transparente de **Subtotal Concreto + Serviço de Bombeamento = Total Geral**.
   - Auditoria de custos ocultos (taxa de bomba, faturamento mínimo, falso frete, estadia de betoneira).
4. **Atualização Contínua de Pacotes & Checklist de Itens Faltantes**:
   - Leitura automática de novas propostas em PDF, imagens (`.jpeg`/`.png`) ou planilhas.
   - Sincronização automática com a planilha Excel de equalização (`.xlsx`).
   - Geração de mensagem resumida para WhatsApp/E-mail com os itens pendentes de cotação por fornecedor.
5. **Simulador de Contraproposta Item a Item & Integração Sienge ERP**:
   - Simulador de desconto (`%`) com tabela detalhada de preços unitários descontados linha a linha e botão para copiar a contraproposta pronta.
   - Aba de exportação formatada com códigos Sienge ERP (`.tsv` / Excel) e impressão nativa em **A4 Paisagem (`@page { size: A4 landscape }`)**.

---

## Como Instalar esta Skill

### Opção 1: Instalação Direta no Antigravity (Windows / macOS / Linux)

Clone ou baixe este repositório diretamente na sua pasta de skills do Antigravity:

**No Windows (PowerShell):**
```powershell
git clone https://github.com/Marcio-Vinicius/mapa-cotacoes.git "$env:USERPROFILE\.gemini\config\skills\mapa-cotacoes"
```

**No macOS / Linux:**
```bash
git clone https://github.com/Marcio-Vinicius/mapa-cotacoes.git ~/.gemini/config/skills/mapa-cotacoes
```

### Opção 2: Via CLI `npx skills`
```bash
npx skills add Marcio-Vinicius/mapa-cotacoes
```

### Opção 3: Download Manual (ZIP)
1. Clique no botão verde **Code** > **Download ZIP** no topo desta página do GitHub.
2. Extraia o arquivo `SKILL.md` dentro da pasta:
   - Windows: `C:\Users\<SEU_USUARIO>\.gemini\config\skills\mapa-cotacoes\SKILL.md`
   - macOS/Linux: `~/.gemini/config/skills/mapa-cotacoes/SKILL.md`

---

## Como Usar no Chat

Após instalar a skill, basta pedir ao agente apontando para a pasta da sua obra:

- **Criar um novo Mapa de Cotações:**
  > *"Use a skill `mapa-cotacoes` na pasta `C:\caminho\da\obra\AÇO` para gerar o mapa de cotações."*
- **Atualizar com novos orçamentos (PDFs ou imagens):**
  > *"Rodar atualização do pacote de aço, atualize o mapa de cotações incluindo os novos preços."*
- **Verificar itens pendentes:**
  > *"Analise para mim e mande uma mensagem resumida de quais itens ainda preciso pedir orçamento para cada fornecedor."*
