# -*- coding: utf-8 -*-
import os
import re
import openpyxl

EXCEL_FILE = r"C:\Users\MarcioGracional\Desktop\ORÇAMENTOS TÉCNICOS\GIARDINI\AÇO\GIA-_-Equalização_Aço_-REV00-_-20260814.xlsx"
TARGET_FILE = r"C:\Users\MarcioGracional\Desktop\ORÇAMENTOS TÉCNICOS\GIARDINI\AÇO\MAPA_DE_COTACAO_GIARDINI_ACO_MODERNO.html"
BRAIN_FILE = r"C:\Users\MarcioGracional\.gemini\antigravity\brain\93a4ffe1-a50f-451d-8fa7-45410697f768\mapa_cotacao_giardini_aco_moderno.html"
LOGO_PATH = r"C:\Users\MarcioGracional\.gemini\antigravity\brain\93a4ffe1-a50f-451d-8fa7-45410697f768\scratch\logo_dark_base64.txt"

def update_excel():
    if not os.path.exists(EXCEL_FILE):
        print(f"Excel file not found: {EXCEL_FILE}")
        return
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["2002_AÇO"]

    # 1. Atualizar Gerdau Barra Reta (Coluna J, linhas 7 a 23) conforme PDFs W09980417 e W10234125 (PT2)
    ws["J7"] = 5.52       # CA-60 5,0mm RT 12M
    ws["J9"] = 5.23       # CA-50 6,3mm RT 12M
    ws["J11"] = 5.23      # CA-50 8,0mm RT 12M
    ws["J13"] = 5.04      # CA-50 10,0mm RT 12M
    ws["J15"] = 4.84      # CA-50 12,5mm RT 12M
    ws["J17"] = 4.84      # CA-50 16,0mm RT 12M
    ws["J19"] = 4.84      # CA-50 20,0mm RT 12M
    ws["J21"] = 4.84      # CA-50 25,0mm RT 12M
    ws["J23"] = 10.62825  # Arame Recozido BWG 16 RL50KG (R$ 10,29 s/ IPI + 3,25% IPI)

    # 2. Atualizar Gerdau Corte e Dobra (Coluna J, linhas 26 a 42) conforme W10234125 (PT2)
    ws["J26"] = 6.44      # CA-60 5,0mm CD
    ws["J28"] = 6.10      # CA-50 6,3mm CD
    ws["J30"] = 6.10      # CA-50 8,0mm CD
    ws["J32"] = 5.82      # CA-50 10,0mm CD
    ws["J34"] = 5.65      # CA-50 12,5mm CD
    ws["J36"] = 5.65      # CA-50 16,0mm CD
    ws["J38"] = 5.65      # CA-50 20,0mm CD
    ws["J40"] = 5.65      # CA-50 25,0mm CD
    ws["J42"] = 10.62825  # Arame Recozido BWG 16 RL50KG

    # 3. Atualizar NORTEFER Barra Reta (Coluna N, linhas 7 a 23) conforme GIA_-_ORÇAMENTO_AÇO_BARRRA_RETA_NORTEFER_REV00.jpeg
    ws["N7"] = 6.05       # CA-60 5,0mm Barra
    ws["N9"] = 5.44       # CA-50 6,3mm Barra
    ws["N11"] = 5.46      # CA-50 8,0mm Barra
    ws["N13"] = 5.33      # CA-50 10,0mm Barra
    ws["N15"] = 5.33      # CA-50 12,5mm Barra
    ws["N17"] = 5.33      # CA-50 16,0mm Barra
    ws["N19"] = 5.23      # CA-50 20,0mm Barra
    ws["N21"] = 5.23      # CA-50 25,0mm Barra
    ws["N23"] = 8.10      # Arame Recozido BWG 16

    # 4. Atualizar NORTEFER Corte e Dobra (Coluna N, linhas 26 a 42) conforme GIA_-_ORÇAMENTO_AÇO_CORTE_E_DOBRA_NORTEFER_REV00.jpeg
    ws["N26"] = 6.05      # CA-60 5,0mm C/D
    ws["N28"] = 6.05      # CA-50 6,3mm C/D
    ws["N30"] = 6.05      # CA-50 8,0mm C/D
    ws["N32"] = 5.85      # CA-50 10,0mm C/D
    ws["N34"] = 5.85      # CA-50 12,5mm C/D
    ws["N36"] = 5.85      # CA-50 16,0mm C/D
    ws["N38"] = 5.85      # CA-50 20,0mm C/D
    ws["N40"] = 5.85      # CA-50 25,0mm C/D
    ws["N42"] = 8.10      # Arame Recozido BWG 16

    # 5. Atualizar Kauefer Barra Reta (Coluna V, linhas 7 a 23) conforme GIA_-_ORÇAMENTO_AÇO_CORTE_E_DOBRA_KAUEFER_REV00.ods
    ws["V7"] = 5.90       # CA-60 5,0mm
    ws["V9"] = 5.49       # CA-50 6,3mm
    ws["V11"] = 5.49      # CA-50 8,0mm
    ws["V13"] = 5.15      # CA-50 10,0mm
    ws["V15"] = 5.15      # CA-50 12,5mm
    ws["V17"] = 5.15      # CA-50 16,0mm
    ws["V19"] = 5.15      # CA-50 20,0mm
    ws["V21"] = 5.15      # CA-50 25,0mm
    ws["V23"] = 11.05     # Arame Recozido BWG 16

    # 6. Atualizar Kauefer Corte e Dobra (Coluna V, linhas 26 a 42) conforme GIA_-_ORÇAMENTO_AÇO_CORTE_E_DOBRA_KAUEFER_REV00.ods
    ws["V26"] = 6.15      # CA-60 5,0mm CD
    ws["V28"] = 6.05      # CA-50 6,3mm CD
    ws["V30"] = 6.05      # CA-50 8,0mm CD
    ws["V32"] = 5.80      # CA-50 10,0mm CD
    ws["V34"] = 5.80      # CA-50 12,5mm CD
    ws["V36"] = 5.80      # CA-50 16,0mm CD
    ws["V38"] = 5.80      # CA-50 20,0mm CD
    ws["V40"] = 5.80      # CA-50 25,0mm CD
    ws["V42"] = 11.05     # Arame Recozido BWG 16

    try:
        wb.save(EXCEL_FILE)
        print(f"Updated Excel successfully: {EXCEL_FILE}")
    except PermissionError:
        alt_excel = EXCEL_FILE.replace(".xlsx", "_ATUALIZADO.xlsx")
        wb.save(alt_excel)
        print(f"Excel was open/locked; saved updated copy to: {alt_excel}")

with open(LOGO_PATH, "r", encoding="utf-8") as f:
    LOGO_BASE64 = f.read().strip()

def build_html():
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Mapa de Cotação de Aço &mdash; Obra 2002 Giardini | Evense</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">

  <style>
    body {{
      font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background-color: #f8fafc;
      color: #0f172a;
      -webkit-font-smoothing: antialiased;
    }}
    .num-tabular {{
      font-variant-numeric: tabular-nums;
    }}
    
    /* Abas de Navegação Estilo SaaS */
    .tab-under {{
      padding: 0.65rem 1.1rem;
      font-size: 0.75rem;
      font-weight: 700;
      color: #64748b;
      border-bottom: 2px solid transparent;
      transition: all 0.15s ease;
      white-space: nowrap;
      display: inline-flex;
      align-items: center;
      gap: 0.5rem;
    }}
    .tab-under:hover {{
      color: #0a2540;
    }}
    .tab-under.active {{
      color: #0a2540;
      border-bottom-color: #0a2540;
      background-color: rgba(10, 37, 64, 0.03);
      border-top-left-radius: 0.5rem;
      border-top-right-radius: 0.5rem;
    }}

    /* Botão Rádio SaaS para Seleção */
    .saas-radio {{
      width: 1.05rem;
      height: 1.05rem;
      border-radius: 9999px;
      border: 1.8px solid #cbd5e1;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      transition: all 0.15s ease;
      cursor: pointer;
      background-color: #ffffff;
      flex-shrink: 0;
    }}
    .saas-radio.checked {{
      border-color: #0a2540;
      background-color: #0a2540;
    }}
    .saas-radio.checked::after {{
      content: '';
      width: 0.38rem;
      height: 0.38rem;
      border-radius: 9999px;
      background-color: #ffffff;
    }}

    /* Células Selecionadas e Menor Preço */
    .cell-selected {{
      background-color: #f0f7ff !important;
    }}
    .cell-best {{
      position: relative;
    }}
    .cell-best::before {{
      content: '';
      position: absolute;
      top: 2px;
      right: 2px;
      width: 5px;
      height: 5px;
      border-radius: 9999px;
      background-color: #10b981;
    }}

    /* Badges de Destaque */
    .diff-pill {{
      font-size: 0.62rem;
      font-weight: 700;
      padding: 0.12rem 0.35rem;
      border-radius: 0.35rem;
      display: inline-block;
    }}
    .diff-pill-best {{
      background-color: #ecfdf5;
      color: #065f46;
      border: 1px solid #a7f3d0;
    }}
    .diff-pill-higher {{
      background-color: #f1f5f9;
      color: #475569;
      border: 1px solid #e2e8f0;
    }}

    /* Tabela Airy com Colunas Fixas */
    .table-airy th, .table-airy td {{
      padding-top: 0.6rem;
      padding-bottom: 0.6rem;
    }}
    .sticky-col-1 {{
      position: sticky;
      left: 0;
      z-index: 20;
    }}
    .sticky-col-2 {{
      position: sticky;
      left: 105px;
      z-index: 20;
    }}

    /* Scrollbar personalizada */
    ::-webkit-scrollbar {{
      width: 6px;
      height: 6px;
    }}
    ::-webkit-scrollbar-track {{
      background: #f1f5f9;
    }}
    ::-webkit-scrollbar-thumb {{
      background: #cbd5e1;
      border-radius: 4px;
    }}
    ::-webkit-scrollbar-thumb:hover {{
      background: #94a3b8;
    }}

    /* Estilos de Impressão A4 Paisagem */
    @media print {{
      @page {{
        size: A4 landscape;
        margin: 5mm;
      }}
      body {{
        background: #ffffff !important;
        font-size: 8pt !important;
        color: #000000 !important;
      }}
      .no-print,
      header.no-print,
      .tab-under,
      .border-b.border-slate-200.flex,
      #export-dropdown,
      #input-search-cd,
      #input-search-br,
      .cursor-pointer input[type="checkbox"],
      .hidden {{
        display: none !important;
      }}
      .print-only {{
        display: block !important;
      }}
      .max-w-\\[1560px\\] {{
        max-width: 100% !important;
        width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
      }}
      .shadow-xs, .shadow-sm, .shadow-lg, .shadow-2xs {{
        box-shadow: none !important;
      }}
      table {{
        width: 100% !important;
        page-break-inside: auto !important;
      }}
      tr {{
        page-break-inside: avoid !important;
        page-break-after: auto !important;
      }}
      tfoot {{
        display: table-row-group !important;
      }}
      tfoot tr {{
        page-break-inside: avoid !important;
      }}
    }}
    .print-only {{
      display: none;
    }}
  </style>
</head>

<body class="p-3 sm:p-5 lg:p-6 min-h-screen flex flex-col justify-between">

  <!-- CABEÇALHO EXCLUSIVO PARA IMPRESSÃO A4 PAISAGEM -->
  <div class="print-only mb-3 border-b-2 border-[#0a2540] pb-2">
    <div class="flex items-center justify-between">
      <div class="flex items-center space-x-3">
        <img src="data:image/png;base64,{LOGO_BASE64}" alt="Evense" style="height: 28px; width: auto;">
        <div>
          <div class="text-xs font-black tracking-wider text-[#0a2540] uppercase">Evense Empreendimentos &bull; Engenharia &amp; Suprimentos</div>
          <div class="text-sm font-bold text-slate-800">Mapa Comparativo de Cotações: Aço &mdash; Obra 2002 Giardini</div>
        </div>
      </div>
      <div class="text-right text-[10px]">
        <div class="font-bold text-slate-800">Eng. Marcio Gracional (Planejamento / Suprimentos)</div>
        <div class="text-slate-500 font-semibold">Atualização Pacote de Aço: 22/10/2026 &bull; REV 03 | Volume Oficial: 18.999,00 kg (19,00 t)</div>
        <div class="text-[9px] text-slate-400">Av. Londrina, 494 - Zona 08, Maringá - PR &bull; Projeto Zocco R01</div>
      </div>
    </div>
  </div>

  <div class="space-y-4 max-w-[1560px] mx-auto w-full">

    <!-- 1. TOP NAVBAR COM O LOGO OFICIAL ESCURO EVENSE -->
    <header class="bg-white rounded-2xl border border-slate-200/80 px-5 py-3 flex flex-wrap items-center justify-between gap-4 shadow-xs no-print">
      <div class="flex items-center space-x-4">
        <div class="flex items-center">
          <img src="data:image/png;base64,{LOGO_BASE64}" alt="Evense Empreendimentos" class="h-9 w-auto object-contain">
        </div>
        <span class="text-slate-300 text-xl font-light">|</span>
        <div>
          <div class="flex items-center space-x-2">
            <span class="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-bold bg-[#0a2540] text-white tracking-wider uppercase">
              Obra 2002 Giardini
            </span>
            <span class="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-semibold bg-blue-50 text-blue-700 border border-blue-200/60">
              REV 03 | 22/10/2026
            </span>
            <span class="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200/60">
              Pacote de Aço Atualizado (9 Propostas)
            </span>
          </div>
          <h1 class="text-lg sm:text-xl font-bold text-slate-900 tracking-tight mt-0.5">
            Mapa de Cotação &mdash; Obra 2002 Giardini
          </h1>
          <p class="text-[11px] text-slate-500">
            Av. Londrina, 494 - Zona 08 &bull; Volume Oficial da Etapa: 18.999,00 kg (18.599,00 kg Aço + 400,00 kg Arame - Projeto Zocco R01)
          </p>
        </div>
      </div>

      <!-- BADGE DO ENGENHEIRO / COMPRADOR & BOTÕES DE AÇÃO -->
      <div class="flex items-center space-x-2.5 flex-wrap gap-y-2">
        <div class="hidden xl:flex items-center space-x-2 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200">
          <div class="w-6 h-6 rounded-full bg-[#0a2540] text-white flex items-center justify-center font-bold text-[10px] shadow-2xs">
            M
          </div>
          <div class="text-left text-xs leading-tight">
            <div class="font-bold text-slate-800">Eng. Marcio Gracional</div>
            <div class="text-[10px] text-slate-500 font-semibold">Eng. de Planejamento / Suprimentos</div>
          </div>
        </div>

        <!-- BOTÃO RESTAURAR VOLUME OFICIAL -->
        <button onclick="resetToOfficialVolume()" class="px-3 py-2 bg-slate-50 hover:bg-slate-100 text-[#0a2540] border border-slate-300/80 rounded-xl text-xs font-bold flex items-center space-x-1.5 shadow-2xs transition" title="Restaura os quantitativos oficiais da Etapa (18.999,00 kg)">
          <svg class="w-3.5 h-3.5 text-[#0a2540]" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M16.023 9.348h4.992v-.001M2.985 19.644v-4.992m0 0h4.992m-4.993 0l3.181 3.183a8.25 8.25 0 0013.803-3.7M4.031 9.865a8.25 8.25 0 0113.803-3.7l3.181 3.182m0-4.991v4.99"/>
          </svg>
          <span>Restaurar 19,00 t</span>
        </button>

        <!-- BOTÃO SALVAR CENÁRIO -->
        <button onclick="saveScenario()" class="px-3 py-2 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300/80 rounded-xl text-xs font-semibold flex items-center space-x-1.5 shadow-2xs transition" title="Salva o cenário atual no navegador">
          <svg class="w-3.5 h-3.5 text-slate-500" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M17.593 3.322c1.1.128 1.907 1.077 1.907 2.185V21L12 17.25 4.5 21V5.507c0-1.108.806-2.057 1.907-2.185a48.507 48.507 0 0111.186 0z"/>
          </svg>
          <span>Salvar cenário</span>
        </button>

        <!-- BOTÃO EXPORTAR COM DROPDOWN -->
        <div class="relative inline-block text-left">
          <button onclick="toggleExportMenu()" class="px-3.5 py-2 bg-[#0a2540] hover:bg-[#12365c] text-white rounded-xl text-xs font-bold flex items-center space-x-2 shadow-sm transition">
            <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3"/>
            </svg>
            <span>Exportar</span>
            <svg class="w-3 h-3 text-slate-300" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M19.5 8.25l-7.5 7.5-7.5-7.5"/>
            </svg>
          </button>

          <!-- MENU DROPDOWN DE EXPORTAÇÃO -->
          <div id="export-dropdown" class="hidden absolute right-0 mt-2 w-52 bg-white border border-slate-200 rounded-xl shadow-lg z-50 p-1.5 text-xs space-y-1">
            <button onclick="exportToXLSX()" class="w-full text-left px-3 py-2 rounded-lg hover:bg-slate-100 font-bold text-slate-800 flex items-center space-x-2">
              <span>Exportar Excel (.xlsx)</span>
            </button>
            <button onclick="exportToCSV()" class="w-full text-left px-3 py-2 rounded-lg hover:bg-slate-100 font-semibold text-slate-700 flex items-center space-x-2">
              <span>Exportar CSV</span>
            </button>
            <button onclick="window.print()" class="w-full text-left px-3 py-2 rounded-lg hover:bg-slate-100 font-semibold text-slate-700 flex items-center space-x-2">
              <span>Imprimir Relatório (PDF)</span>
            </button>
          </div>
        </div>
      </div>
    </header>

    <!-- 2. TOP METRIC CARDS (5 KPIS REAIS E DINÂMICOS) -->
    <div id="top-kpi-container" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
      
      <!-- CARD 1: Custo total selecionado -->
      <div class="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-xs flex items-center space-x-3.5">
        <div class="w-11 h-11 rounded-2xl bg-blue-50 text-blue-600 flex items-center justify-center flex-shrink-0">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="1.8" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M20.25 6.375c0 2.278-3.694 4.125-8.25 4.125S3.75 8.653 3.75 6.375m16.5 0c0-2.278-3.694-4.125-8.25-4.125S3.75 4.097 3.75 6.375m16.5 0v11.25c0 2.278-3.694 4.125-8.25 4.125s-8.25-1.847-8.25-4.125V6.375m16.5 5.625c0 2.278-3.694 4.125-8.25 4.125s-8.25-1.847-8.25-4.125"/>
          </svg>
        </div>
        <div>
          <div class="text-[11px] font-semibold text-slate-500">Custo total selecionado</div>
          <div class="text-lg font-extrabold text-[#0a2540] tracking-tight mt-0.5 num-tabular" id="card-custo-total">R$ 104.687,64</div>
          <div class="text-[10px] font-bold text-emerald-600 flex items-center space-x-1 mt-0.5" id="card-custo-diff">
            <span>&darr;</span>
            <span>-7,7% vs. pior proposta</span>
          </div>
        </div>
      </div>

      <!-- CARD 2: Economia potencial -->
      <div class="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-xs flex items-center space-x-3.5">
        <div class="w-11 h-11 rounded-2xl bg-emerald-50 text-emerald-600 flex items-center justify-center flex-shrink-0">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="1.8" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M3 13.125C3 12.504 3.504 12 4.125 12h2.25c.621 0 1.125.504 1.125 1.125v6.75C7.5 20.496 6.996 21 6.375 21h-2.25A1.125 1.125 0 013 19.875v-6.75zM9.75 8.625c0-.621.504-1.125 1.125-1.125h2.25c.621 0 1.125.504 1.125 1.125v11.25c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 01-1.125-1.125V8.625zM16.5 4.125c0-.621.504-1.125 1.125-1.125h2.25C20.496 3 21 3.504 21 4.125v15.75c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 01-1.125-1.125V4.125z"/>
          </svg>
        </div>
        <div>
          <div class="text-[11px] font-semibold text-slate-500">Economia potencial</div>
          <div class="text-lg font-extrabold text-[#0a2540] tracking-tight mt-0.5 num-tabular" id="card-economia-potencial">R$ 8.698,66</div>
          <div class="text-[10px] font-bold text-emerald-600 mt-0.5" id="card-economia-pct">7,7% de redução vs. maior CD</div>
        </div>
      </div>

      <!-- CARD 3: Volume total -->
      <div class="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-xs flex items-center space-x-3.5">
        <div class="w-11 h-11 rounded-2xl bg-indigo-50 text-indigo-600 flex items-center justify-center flex-shrink-0">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="1.8" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M21 7.5l-9-5.25L3 7.5m18 0l-9 5.25m9-5.25v9l-9 5.25M3 7.5l9 5.25M3 7.5v9l9 5.25m0-9v9"/>
          </svg>
        </div>
        <div>
          <div class="text-[11px] font-semibold text-slate-500">Volume em cotação</div>
          <div class="text-lg font-extrabold text-[#0a2540] tracking-tight mt-0.5 num-tabular" id="card-volume-total">18.999,00 kg</div>
          <div class="text-[10px] text-slate-400 mt-0.5" id="card-itens-count">5 itens com volume ativo</div>
        </div>
      </div>

      <!-- CARD 4: Fornecedores na seleção -->
      <div class="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-xs flex items-center space-x-3.5">
        <div class="w-11 h-11 rounded-2xl bg-sky-50 text-sky-600 flex items-center justify-center flex-shrink-0">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="1.8" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M8.25 18.75a1.5 1.5 0 01-3 0m3 0a1.5 1.5 0 00-3 0m3 0h6m-9 0H3.375a1.125 1.125 0 01-1.125-1.125V14.25m17.25 4.5a1.5 1.5 0 01-3 0m3 0a1.5 1.5 0 00-3 0m3 0h1.125c.621 0 1.129-.504 1.09-1.124a17.902 17.902 0 00-3.213-9.193 2.056 2.056 0 00-1.58-.86H14.25M16.5 18.75h-2.25m0-11.25h2.25m-2.25 0V3.75"/>
          </svg>
        </div>
        <div>
          <div class="text-[11px] font-semibold text-slate-500">Fornecedores na seleção</div>
          <div class="text-lg font-extrabold text-[#0a2540] tracking-tight mt-0.5 num-tabular" id="card-fornecedores-count">2</div>
          <div class="text-[10px] text-slate-400 mt-0.5" id="card-fornecedores-selecionados">Mix (Paranafer + Nortefer)</div>
        </div>
      </div>

      <!-- CARD 5: Cobertura da proposta -->
      <div class="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-xs flex items-center space-x-3.5">
        <div class="w-11 h-11 rounded-2xl bg-emerald-50 text-emerald-600 flex items-center justify-center flex-shrink-0">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="1.8" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12c0 1.268-.63 2.39-1.593 3.068a3.745 3.745 0 01-1.043 3.296 3.745 3.745 0 01-3.296 1.043A3.745 3.745 0 0112 21c-1.268 0-2.39-.63-3.068-1.593a3.746 3.746 0 01-3.296-1.043 3.745 3.745 0 01-1.043-3.296A3.745 3.745 0 013 12c0-1.268.63-2.39 1.593-3.068a3.745 3.745 0 011.043-3.296 3.746 3.746 0 013.296-1.043A3.746 3.746 0 0112 3c1.268 0 2.39.63 3.068 1.593a3.746 3.746 0 013.296 1.043 3.746 3.746 0 011.043 3.296A3.745 3.745 0 0121 12z"/>
          </svg>
        </div>
        <div>
          <div class="text-[11px] font-semibold text-slate-500">Status do mapa</div>
          <div class="text-lg font-extrabold text-emerald-700 tracking-tight mt-0.5" id="card-cobertura-pct">100% Equalizado</div>
          <div class="text-[10px] text-slate-400 mt-0.5" id="card-status-mapa">5 cotações CD + 4 Barra</div>
        </div>
      </div>

    </div>

    <!-- 3. ABAS DE NAVEGAÇÃO MODERNAS (2 ABAS COMPARATIVAS INDEPENDENTES) -->
    <div class="border-b border-slate-200 flex space-x-1 overflow-x-auto no-print">
      <button onclick="switchTab('corte-dobra')" id="tab-btn-corte-dobra" class="tab-under active">
        <span>Corte e Dobra Industrial (5)</span>
      </button>
      <button onclick="switchTab('barra-reta')" id="tab-btn-barra-reta" class="tab-under">
        <span>Barra Reta 12m (4)</span>
      </button>
      <button onclick="switchTab('resumo')" id="tab-btn-resumo" class="tab-under">
        <span>Resumo Executivo &amp; Ranking Geral</span>
      </button>
      <button onclick="switchTab('simulacoes')" id="tab-btn-simulacoes" class="tab-under">
        <span>Simulações &amp; Negociação</span>
      </button>
      <button onclick="switchTab('equalizacao')" id="tab-btn-equalizacao" class="tab-under">
        <span>Equalização Técnica (CD vs. Barra)</span>
      </button>
      <button onclick="switchTab('tecnica')" id="tab-btn-tecnica" class="tab-under">
        <span>Análise Técnica &amp; Normas</span>
      </button>
      <button onclick="switchTab('sienge')" id="tab-btn-sienge" class="tab-under">
        <span>Importação Sienge ERP</span>
      </button>
      <button onclick="switchTab('graficos')" id="tab-btn-graficos" class="tab-under">
        <span>Gráficos Comparativos</span>
      </button>
    </div>

    <!-- ==================== ABA 1: CORTE E DOBRA INDUSTRIAL (5 FORNECEDORES) ==================== -->
    <div id="tab-content-corte-dobra" class="space-y-5">
      
      <!-- TOOLBAR DE CONTROLES CD -->
      <div class="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-xs flex flex-wrap items-center justify-between gap-3 no-print">
        <div>
          <h3 class="text-sm font-bold text-[#0a2540]">Mapa Comparativo: Corte e Dobra Industrial (5 Fornecedores)</h3>
          <p class="text-xs text-slate-500">Cotação atualizada para peças etiquetadas e cortadas conforme projeto estrutural Zocco R01</p>
        </div>

        <div class="flex items-center space-x-2 flex-wrap gap-y-2">
          <!-- BUSCA -->
          <div class="relative">
            <input type="text" id="input-search-cd" oninput="handleSearchCD(this.value)" placeholder="Buscar bitola ou código..." class="pl-8 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-1 focus:ring-[#0a2540] w-44 transition">
            <svg class="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607z"/>
            </svg>
          </div>

          <!-- FILTRO CATEGORIA -->
          <select id="filter-categoria-cd" onchange="setFilterCD(this.value)" class="py-1.5 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none font-semibold text-slate-700">
            <option value="all">Todas Categorias (9 itens)</option>
            <option value="ca50">Aço CA-50 (7 bitolas)</option>
            <option value="ca60">Aço CA-60 (1 bitola)</option>
            <option value="arame">Arame Recozido (1 item)</option>
          </select>

          <!-- TOGGLE APENAS COM VOLUME -->
          <label class="flex items-center space-x-1.5 text-xs font-semibold text-slate-600 cursor-pointer bg-slate-50 px-2.5 py-1.5 rounded-xl border border-slate-200 hover:bg-slate-100 transition">
            <input type="checkbox" id="toggle-active-cd" onchange="toggleOnlyActiveCD(this.checked)" class="rounded text-[#0a2540] focus:ring-[#0a2540]">
            <span>Apenas com volume (&gt;0 kg)</span>
          </label>

          <!-- BOTÕES DE SELEÇÃO RÁPIDA (5 FORNECEDORES CD) -->
          <button onclick="selectSingleSupplierCD(0)" class="px-2.5 py-1.5 bg-cyan-50 hover:bg-cyan-100 text-cyan-900 border border-cyan-200/80 rounded-xl text-xs font-bold transition">
            Paranafer (1º Aço CD)
          </button>
          <button onclick="selectSingleSupplierCD(1)" class="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200 rounded-xl text-xs font-bold transition">
            Arcelor CD (1º Compl.)
          </button>
          <button onclick="selectSingleSupplierCD(2)" class="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200 rounded-xl text-xs font-bold transition">
            Gerdau CD (2º Compl.)
          </button>
          <button onclick="selectSingleSupplierCD(3)" class="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200 rounded-xl text-xs font-bold transition">
            Kauefer CD (3º Compl.)
          </button>
          <button onclick="selectSingleSupplierCD(4)" class="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200 rounded-xl text-xs font-bold transition">
            NORTEFER CD (4º Compl.)
          </button>
          <button onclick="selectBestPricesAllCD()" class="px-2.5 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 rounded-xl text-xs font-bold transition">
            Menor Item a Item
          </button>
        </div>
      </div>

      <!-- TABELA CORTE E DOBRA (5 FORNECEDORES) -->
      <div class="bg-white rounded-2xl border border-slate-200/80 shadow-xs overflow-hidden">
        <div class="overflow-x-auto">
          <table class="w-full text-left text-xs border-collapse table-airy" id="table-cd">
            <thead>
              <tr class="bg-slate-50/90 border-b border-slate-200 text-slate-700 font-bold">
                <th class="py-3 px-2.5 w-24 sticky-col-1 bg-slate-50">Código</th>
                <th class="py-3 px-2.5 min-w-[205px] sticky-col-2 bg-slate-50">Descrição do Insumo</th>
                <th class="py-3 px-1.5 text-center w-12">Unid.</th>
                <th class="py-3 px-2.5 text-right w-24">Volume (kg)</th>

                <!-- FORNECEDOR 1 CD: PARANAFER -->
                <th class="py-3 px-2.5 min-w-[168px] border-l border-slate-200 bg-cyan-50/40 text-center cursor-pointer hover:bg-cyan-50 transition" onclick="selectSingleSupplierCD(0)">
                  <div class="flex items-center justify-center space-x-1.5">
                    <span class="font-extrabold text-cyan-950 text-xs sm:text-sm">Paranafer</span>
                    <span class="text-[9px] bg-cyan-200 text-cyan-900 font-black px-1.5 py-0.5 rounded">1º Aço CD</span>
                  </div>
                  <div class="text-[10px] text-slate-500 font-normal mt-0.5">Orç. 70.699 &bull; 30D Boleto (S/ Arame)</div>
                </th>

                <!-- FORNECEDOR 2 CD: ARCELORMITTAL CD -->
                <th class="py-3 px-2.5 min-w-[168px] border-l border-slate-200 text-center cursor-pointer hover:bg-slate-100 transition" onclick="selectSingleSupplierCD(1)">
                  <div class="flex items-center justify-center space-x-1.5">
                    <span class="font-extrabold text-slate-900 text-xs sm:text-sm">ArcelorMittal (CD)</span>
                    <span class="text-[9px] bg-emerald-100 text-emerald-800 font-black px-1.5 py-0.5 rounded">1º Compl.</span>
                  </div>
                  <div class="text-[10px] text-slate-500 font-normal mt-0.5">Belgo Pronto &bull; 15 Dias CIF</div>
                </th>

                <!-- FORNECEDOR 3 CD: GERDAU CD -->
                <th class="py-3 px-2.5 min-w-[168px] border-l border-slate-200 text-center cursor-pointer hover:bg-slate-100 transition" onclick="selectSingleSupplierCD(2)">
                  <div class="flex items-center justify-center space-x-1.5">
                    <span class="font-extrabold text-slate-900 text-xs sm:text-sm">Gerdau (CD)</span>
                    <span class="text-[9px] bg-slate-200 text-slate-700 font-bold px-1.5 py-0.5 rounded">2º Compl.</span>
                  </div>
                  <div class="text-[10px] text-slate-500 font-normal mt-0.5">Cotação W10234125 &bull; Usina CIF</div>
                </th>

                <!-- FORNECEDOR 4 CD: KAUEFER CD -->
                <th class="py-3 px-2.5 min-w-[168px] border-l border-slate-200 text-center cursor-pointer hover:bg-slate-100 transition" onclick="selectSingleSupplierCD(3)">
                  <div class="flex items-center justify-center space-x-1.5">
                    <span class="font-extrabold text-slate-900 text-xs sm:text-sm">Kauefer (CD)</span>
                    <span class="text-[9px] bg-slate-200 text-slate-700 font-bold px-1.5 py-0.5 rounded">3º Compl.</span>
                  </div>
                  <div class="text-[10px] text-slate-500 font-normal mt-0.5">Corte e Dobra &bull; Faturado CIF</div>
                </th>

                <!-- FORNECEDOR 5 CD: NORTEFER CD -->
                <th class="py-3 px-2.5 min-w-[168px] border-l border-slate-200 text-center cursor-pointer hover:bg-slate-100 transition" onclick="selectSingleSupplierCD(4)">
                  <div class="flex items-center justify-center space-x-1.5">
                    <span class="font-extrabold text-slate-900 text-xs sm:text-sm">NORTEFER (CD)</span>
                    <span class="text-[9px] bg-slate-200 text-slate-700 font-bold px-1.5 py-0.5 rounded">4º Compl.</span>
                  </div>
                  <div class="text-[10px] text-slate-500 font-normal mt-0.5">Corte e Dobra &bull; Atualizado 22/10</div>
                </th>
              </tr>
            </thead>
            <tbody id="table-body-cd" class="divide-y divide-slate-100 text-slate-800">
              <!-- Gerado via JS -->
            </tbody>
            <tfoot class="border-t-2 border-slate-200 font-bold bg-slate-50/70">
              <!-- LINHA 1: SUBTOTAL AÇO -->
              <tr class="text-slate-700">
                <td colspan="4" class="py-2.5 px-3 text-right font-extrabold sticky-col-1 bg-slate-50">
                  Subtotal Aço CA-50 / CA-60 (18.599,00 kg):
                </td>
                <td class="py-2.5 px-2.5 text-right border-l border-slate-200 num-tabular" id="subtot-cd-0">R$ 101.447,64</td>
                <td class="py-2.5 px-2.5 text-right border-l border-slate-200 num-tabular" id="subtot-cd-1">R$ 105.794,45</td>
                <td class="py-2.5 px-2.5 text-right border-l border-slate-200 num-tabular" id="subtot-cd-2">R$ 106.505,14</td>
                <td class="py-2.5 px-2.5 text-right border-l border-slate-200 num-tabular" id="subtot-cd-3">R$ 108.464,65</td>
                <td class="py-2.5 px-2.5 text-right border-l border-slate-200 num-tabular" id="subtot-cd-4">R$ 109.141,55</td>
              </tr>
              <!-- LINHA 2: ARAME RECOZIDO -->
              <tr class="text-slate-600 text-xs">
                <td colspan="4" class="py-2 px-3 text-right font-semibold sticky-col-1 bg-slate-50">
                  Arame Recozido BWG 16 (400,00 kg):
                </td>
                <td class="py-2 px-2.5 text-right border-l border-slate-200 num-tabular text-amber-700 italic" id="arame-cd-0">Não cotado</td>
                <td class="py-2 px-2.5 text-right border-l border-slate-200 num-tabular text-slate-600" id="arame-cd-1">R$ 3.244,00</td>
                <td class="py-2 px-2.5 text-right border-l border-slate-200 num-tabular text-slate-600" id="arame-cd-2">R$ 4.251,30</td>
                <td class="py-2 px-2.5 text-right border-l border-slate-200 num-tabular text-slate-600" id="arame-cd-3">R$ 4.420,00</td>
                <td class="py-2 px-2.5 text-right border-l border-slate-200 num-tabular text-slate-600" id="arame-cd-4">R$ 3.240,00</td>
              </tr>
              <!-- LINHA 3: TOTAL GERAL DA PROPOSTA -->
              <tr class="bg-slate-100 text-slate-900 text-sm">
                <td colspan="4" class="py-3 px-3 text-right font-black sticky-col-1 bg-slate-100">
                  Total Geral Cotado na Proposta:
                </td>
                <td class="py-3 px-2.5 text-right border-l border-slate-200 font-black text-cyan-950 bg-cyan-100/40 num-tabular" id="tot-cd-0">R$ 101.447,64</td>
                <td class="py-3 px-2.5 text-right border-l border-slate-200 font-black text-slate-900 num-tabular" id="tot-cd-1">R$ 109.038,45</td>
                <td class="py-3 px-2.5 text-right border-l border-slate-200 font-black text-slate-900 num-tabular" id="tot-cd-2">R$ 110.756,44</td>
                <td class="py-3 px-2.5 text-right border-l border-slate-200 font-black text-slate-900 num-tabular" id="tot-cd-3">R$ 112.884,65</td>
                <td class="py-3 px-2.5 text-right border-l border-slate-200 font-black text-slate-900 num-tabular" id="tot-cd-4">R$ 112.381,55</td>
              </tr>
              <!-- LINHA 4: PRAZOS E CONDIÇÃO COMERCIAL -->
              <tr class="text-[11px] text-slate-600 bg-white">
                <td colspan="4" class="py-2 px-3 text-right font-semibold sticky-col-1 bg-white">Condição de Pagamento / Frete:</td>
                <td class="py-2 px-2.5 text-center border-l border-slate-200 text-cyan-900 font-bold">30 Dias Boleto &bull; CIF</td>
                <td class="py-2 px-2.5 text-center border-l border-slate-200 font-semibold">15 Dias &bull; CIF</td>
                <td class="py-2 px-2.5 text-center border-l border-slate-200 font-semibold">Antecipado B014 &bull; CIF</td>
                <td class="py-2 px-2.5 text-center border-l border-slate-200 font-semibold">Faturado &bull; CIF</td>
                <td class="py-2 px-2.5 text-center border-l border-slate-200 font-semibold">Faturado Programado &bull; CIF</td>
              </tr>
              <!-- LINHA 5: STATUS E CLASSIFICAÇÃO -->
              <tr class="text-xs bg-[#0a2540] text-white">
                <td colspan="4" class="py-2.5 px-3 text-right font-bold sticky-col-1 bg-[#0a2540]">Classificação em Corte e Dobra:</td>
                <td class="py-2.5 px-2.5 text-center border-l border-slate-700" id="stat-cd-0"><span class="text-amber-300 font-black">1º no Aço CD (S/ Arame)</span></td>
                <td class="py-2.5 px-2.5 text-center border-l border-slate-700" id="stat-cd-1"><span class="text-emerald-300 font-black">1º Completo (c/ Arame)</span></td>
                <td class="py-2.5 px-2.5 text-center border-l border-slate-700" id="stat-cd-2"><span>2º Completo (+1,6%)</span></td>
                <td class="py-2.5 px-2.5 text-center border-l border-slate-700" id="stat-cd-3"><span>4º Completo (+3,5%)</span></td>
                <td class="py-2.5 px-2.5 text-center border-l border-slate-700" id="stat-cd-4"><span>3º Completo (+3,1%)</span></td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      <!-- BOTTOM DASHBOARD SECTION CD (3 CARDS LADO A LADO) -->
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <!-- CARD 1: DONUT CHART CD -->
        <div class="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div>
            <div class="flex items-center justify-between border-b border-slate-100 pb-2">
              <h4 class="font-bold text-xs text-[#0a2540] uppercase tracking-wider">Distribuição do Custo (CD)</h4>
              <span class="text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded font-bold">Corte e Dobra</span>
            </div>
            <div class="flex items-center justify-center my-3">
              <div class="relative w-36 h-36 flex items-center justify-center">
                <svg id="donut-svg-cd" class="w-full h-full transform -rotate-90" viewBox="0 0 100 100">
                  <circle cx="50" cy="50" r="38" stroke="#f1f5f9" stroke-width="14" fill="transparent" />
                </svg>
                <div class="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                  <span class="text-xs font-black text-slate-900 num-tabular" id="donut-center-val-cd">R$ 105 mil</span>
                  <span class="text-[9px] text-slate-400 font-semibold">Total CD</span>
                </div>
              </div>
            </div>
          </div>
          <div id="donut-legend-cd" class="space-y-1.5 text-xs pt-1 border-t border-slate-100">
            <!-- Gerado via JS -->
          </div>
        </div>

        <!-- CARD 2: COMPOSIÇÃO DE BITOLAS & INSPEÇÃO CD -->
        <div class="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div>
            <div class="flex items-center justify-between border-b border-slate-100 pb-2">
              <h4 class="font-bold text-xs text-[#0a2540] uppercase tracking-wider">Composição de Bitolas &amp; Item</h4>
              <span class="text-[10px] text-slate-500 font-semibold">18.999,00 kg total</span>
            </div>
            
            <div class="mt-3">
              <div class="text-[11px] font-semibold text-slate-600 mb-1 flex justify-between">
                <span>Distribuição do Peso por Bitola</span>
                <span class="text-slate-400">12,5mm = 66,7%</span>
              </div>
              <div class="h-3 w-full bg-slate-100 rounded-full overflow-hidden flex" id="composition-bar-cd">
                <!-- Gerado via JS -->
              </div>
              <div class="flex justify-between text-[9px] text-slate-400 mt-1" id="composition-legend-cd">
                <!-- Gerado via JS -->
              </div>
            </div>

            <!-- DETALHE TÉCNICO DO ITEM INSPECIONADO -->
            <div class="mt-3.5 p-3 bg-slate-50 rounded-xl border border-slate-200/80" id="detail-card-cd">
              <div class="flex items-center justify-between mb-1.5">
                <span class="text-[10px] font-bold text-slate-500 uppercase" id="detail-item-cod-cd">01.002.0012</span>
                <span class="text-[10px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-800 font-bold" id="detail-item-cat-cd">CA-50</span>
              </div>
              <div class="font-bold text-xs text-slate-900" id="detail-item-name-cd">FERRO CA-50 12,5MM (1/2")</div>
              <div class="text-[11px] text-slate-500 mt-0.5" id="detail-item-desc-cd">Armadura principal de vigas e pilares mestres</div>
              
              <div class="grid grid-cols-3 gap-2 mt-2.5 pt-2 border-t border-slate-200 text-center text-xs">
                <div>
                  <div class="text-[9px] text-slate-400">Volume</div>
                  <div class="font-bold text-slate-800" id="detail-item-vol-cd">12.672,00 kg</div>
                </div>
                <div>
                  <div class="text-[9px] text-slate-400">Preço Sel.</div>
                  <div class="font-bold text-cyan-900" id="detail-item-price-cd">R$ 5,42/kg</div>
                </div>
                <div>
                  <div class="text-[9px] text-slate-400">Total Item</div>
                  <div class="font-bold text-slate-900" id="detail-item-total-cd">R$ 68.682,24</div>
                </div>
              </div>
            </div>
          </div>
          <div class="text-[10px] text-slate-400 text-center mt-2">Clique em qualquer linha da tabela para inspecionar a bitola</div>
        </div>

        <!-- CARD 3: RESUMO DO CENÁRIO CORTE E DOBRA -->
        <div class="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div>
            <div class="flex items-center justify-between border-b border-slate-100 pb-2">
              <h4 class="font-bold text-xs text-[#0a2540] uppercase tracking-wider">Resumo do Cenário (CD)</h4>
              <span class="text-[10px] bg-cyan-100 text-cyan-800 px-2 py-0.5 rounded font-bold">100% Corte e Dobra</span>
            </div>

            <div class="space-y-2.5 mt-3 text-xs">
              <div class="flex justify-between items-center py-1 border-b border-slate-100">
                <span class="text-slate-500">Custo Total CD Selecionado:</span>
                <span class="font-black text-slate-900 num-tabular text-sm" id="summary-custo-total-cd">R$ 104.687,64</span>
              </div>
              <div class="flex justify-between items-center py-1 border-b border-slate-100">
                <span class="text-slate-500">Menor Mix Item a Item (CD):</span>
                <span class="font-bold text-cyan-800 num-tabular" id="summary-custo-menor-cd">R$ 104.687,64</span>
              </div>
              <div class="flex justify-between items-center py-1 border-b border-slate-100">
                <span class="text-slate-500">1º Fornecedor Único Completo:</span>
                <span class="font-bold text-slate-800 num-tabular">Arcelor CD (R$ 109.038,45)</span>
              </div>
              <div class="flex justify-between items-center py-1 border-b border-slate-100">
                <span class="text-slate-500">Economia vs. Kauefer CD:</span>
                <span class="font-bold text-emerald-700 num-tabular" id="summary-economia-cd">R$ 8.197,01 (-7,3%)</span>
              </div>
              <div class="flex justify-between items-center py-1 border-b border-slate-100">
                <span class="text-slate-500">Condição Comercial Recomendada:</span>
                <span class="font-bold text-slate-800">30D Boleto CIF (Paranafer)</span>
              </div>
              <div class="flex justify-between items-center py-1">
                <span class="text-slate-500">Status de Aprovação:</span>
                <span class="font-bold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded">Pronto p/ Homologação</span>
              </div>
            </div>
          </div>

          <button onclick="switchTab('simulacoes')" class="w-full mt-3 py-2 bg-slate-50 hover:bg-slate-100 text-[#0a2540] border border-slate-200 rounded-xl text-xs font-bold transition flex items-center justify-center space-x-1.5">
            <span>Abrir Simulador de Contraproposta</span>
            <span>&rarr;</span>
          </button>
        </div>
      </div>

    </div>

    <!-- ==================== ABA 2: BARRA RETA (12M) (4 FORNECEDORES) ==================== -->
    <div id="tab-content-barra-reta" class="space-y-5 hidden">
      
      <!-- TOOLBAR DE CONTROLES BARRA RETA -->
      <div class="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-xs flex flex-wrap items-center justify-between gap-3 no-print">
        <div>
          <h3 class="text-sm font-bold text-[#0a2540]">Mapa Comparativo: Barra Reta (12 metros) &mdash; 4 Fornecedores</h3>
          <p class="text-xs text-slate-500">Cotação atualizada com todas as 9 bitolas preenchidas nos 4 fornecedores de Barra Reta (incluindo Nortefer)</p>
        </div>

        <div class="flex items-center space-x-2 flex-wrap gap-y-2">
          <!-- BUSCA -->
          <div class="relative">
            <input type="text" id="input-search-br" oninput="handleSearchBR(this.value)" placeholder="Buscar bitola ou código..." class="pl-8 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-1 focus:ring-[#0a2540] w-44 transition">
            <svg class="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607z"/>
            </svg>
          </div>

          <!-- FILTRO CATEGORIA -->
          <select id="filter-categoria-br" onchange="setFilterBR(this.value)" class="py-1.5 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none font-semibold text-slate-700">
            <option value="all">Todas Categorias (9 itens)</option>
            <option value="ca50">Aço CA-50 (7 bitolas)</option>
            <option value="ca60">Aço CA-60 (1 bitola)</option>
            <option value="arame">Arame Recozido (1 item)</option>
          </select>

          <!-- TOGGLE APENAS COM VOLUME -->
          <label class="flex items-center space-x-1.5 text-xs font-semibold text-slate-600 cursor-pointer bg-slate-50 px-2.5 py-1.5 rounded-xl border border-slate-200 hover:bg-slate-100 transition">
            <input type="checkbox" id="toggle-active-br" onchange="toggleOnlyActiveBR(this.checked)" class="rounded text-[#0a2540] focus:ring-[#0a2540]">
            <span>Apenas com volume (&gt;0 kg)</span>
          </label>

          <!-- BOTÕES DE SELEÇÃO RÁPIDA (4 FORNECEDORES BARRA) -->
          <button onclick="selectSingleSupplierBR(0)" class="px-2.5 py-1.5 bg-blue-50 hover:bg-blue-100 text-blue-900 border border-blue-200/80 rounded-xl text-xs font-bold transition">
            ArcelorMittal (1º Barra)
          </button>
          <button onclick="selectSingleSupplierBR(1)" class="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200 rounded-xl text-xs font-bold transition">
            Gerdau (2º Barra)
          </button>
          <button onclick="selectSingleSupplierBR(2)" class="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200 rounded-xl text-xs font-bold transition">
            Kauefer (3º Barra)
          </button>
          <button onclick="selectSingleSupplierBR(3)" class="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200 rounded-xl text-xs font-bold transition">
            NORTEFER (4º Barra)
          </button>
          <button onclick="selectBestPricesAllBR()" class="px-2.5 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 rounded-xl text-xs font-bold transition">
            Menor Item a Item
          </button>
        </div>
      </div>

      <!-- TABELA BARRA RETA (4 FORNECEDORES) -->
      <div class="bg-white rounded-2xl border border-slate-200/80 shadow-xs overflow-hidden">
        <div class="overflow-x-auto">
          <table class="w-full text-left text-xs border-collapse table-airy" id="table-br">
            <thead>
              <tr class="bg-slate-50/90 border-b border-slate-200 text-slate-700 font-bold">
                <th class="py-3 px-2.5 w-24 sticky-col-1 bg-slate-50">Código</th>
                <th class="py-3 px-2.5 min-w-[210px] sticky-col-2 bg-slate-50">Descrição do Insumo</th>
                <th class="py-3 px-1.5 text-center w-12">Unid.</th>
                <th class="py-3 px-2.5 text-right w-24">Volume (kg)</th>

                <!-- FORNECEDOR 1 BARRA: ARCELORMITTAL -->
                <th class="py-3 px-2.5 min-w-[180px] border-l border-slate-200 bg-blue-50/40 text-center cursor-pointer hover:bg-blue-50 transition" onclick="selectSingleSupplierBR(0)">
                  <div class="flex items-center justify-center space-x-1.5">
                    <span class="font-extrabold text-blue-950 text-xs sm:text-sm">ArcelorMittal</span>
                    <span class="text-[9px] bg-blue-200 text-blue-900 font-black px-1.5 py-0.5 rounded">1º Barra</span>
                  </div>
                  <div class="text-[10px] text-slate-500 font-normal mt-0.5">Barra Reta (12m) &bull; Mary 44 9916-9764</div>
                </th>

                <!-- FORNECEDOR 2 BARRA: GERDAU -->
                <th class="py-3 px-2.5 min-w-[180px] border-l border-slate-200 text-center cursor-pointer hover:bg-slate-100 transition" onclick="selectSingleSupplierBR(1)">
                  <div class="flex items-center justify-center space-x-1.5">
                    <span class="font-extrabold text-slate-900 text-xs sm:text-sm">Gerdau</span>
                    <span class="text-[9px] bg-slate-200 text-slate-700 font-bold px-1.5 py-0.5 rounded">2º Barra</span>
                  </div>
                  <div class="text-[10px] text-slate-500 font-normal mt-0.5">W09980417 / W10234125 &bull; CIF</div>
                </th>

                <!-- FORNECEDOR 3 BARRA: KAUEFER -->
                <th class="py-3 px-2.5 min-w-[180px] border-l border-slate-200 text-center cursor-pointer hover:bg-slate-100 transition" onclick="selectSingleSupplierBR(2)">
                  <div class="flex items-center justify-center space-x-1.5">
                    <span class="font-extrabold text-slate-900 text-xs sm:text-sm">Kauefer</span>
                    <span class="text-[9px] bg-slate-200 text-slate-700 font-bold px-1.5 py-0.5 rounded">3º Barra</span>
                  </div>
                  <div class="text-[10px] text-slate-500 font-normal mt-0.5">Barra Reta (12m) &bull; Distribuidor CIF</div>
                </th>

                <!-- FORNECEDOR 4 BARRA: NORTEFER -->
                <th class="py-3 px-2.5 min-w-[180px] border-l border-slate-200 text-center cursor-pointer hover:bg-slate-100 transition" onclick="selectSingleSupplierBR(3)">
                  <div class="flex items-center justify-center space-x-1.5">
                    <span class="font-extrabold text-slate-900 text-xs sm:text-sm">NORTEFER</span>
                    <span class="text-[9px] bg-slate-200 text-slate-700 font-bold px-1.5 py-0.5 rounded">4º Barra</span>
                  </div>
                  <div class="text-[10px] text-slate-500 font-normal mt-0.5">Barra Corte Sob Medida &bull; CIF</div>
                </th>
              </tr>
            </thead>
            <tbody id="table-body-br" class="divide-y divide-slate-100 text-slate-800">
              <!-- Gerado via JS -->
            </tbody>
            <tfoot class="border-t-2 border-slate-200 font-bold bg-slate-50/70">
              <!-- LINHA 1: SUBTOTAL AÇO -->
              <tr class="text-slate-700">
                <td colspan="4" class="py-2.5 px-3 text-right font-extrabold sticky-col-1 bg-slate-50">
                  Subtotal Aço CA-50 / CA-60 (18.599,00 kg):
                </td>
                <td class="py-2.5 px-2.5 text-right border-l border-slate-200 num-tabular" id="subtot-br-0">R$ 90.543,27</td>
                <td class="py-2.5 px-2.5 text-right border-l border-slate-200 num-tabular" id="subtot-br-1">R$ 91.269,92</td>
                <td class="py-2.5 px-2.5 text-right border-l border-slate-200 num-tabular" id="subtot-br-2">R$ 97.050,10</td>
                <td class="py-2.5 px-2.5 text-right border-l border-slate-200 num-tabular" id="subtot-br-3">R$ 100.347,31</td>
              </tr>
              <!-- LINHA 2: ARAME RECOZIDO -->
              <tr class="text-slate-600 text-xs">
                <td colspan="4" class="py-2 px-3 text-right font-semibold sticky-col-1 bg-slate-50">
                  Arame Recozido BWG 16 (400,00 kg):
                </td>
                <td class="py-2 px-2.5 text-right border-l border-slate-200 num-tabular text-slate-600" id="arame-br-0">R$ 3.244,00</td>
                <td class="py-2 px-2.5 text-right border-l border-slate-200 num-tabular text-slate-600" id="arame-br-1">R$ 4.251,30</td>
                <td class="py-2 px-2.5 text-right border-l border-slate-200 num-tabular text-slate-600" id="arame-br-2">R$ 4.420,00</td>
                <td class="py-2 px-2.5 text-right border-l border-slate-200 num-tabular text-slate-600" id="arame-br-3">R$ 3.240,00</td>
              </tr>
              <!-- LINHA 3: TOTAL GERAL DA PROPOSTA -->
              <tr class="bg-slate-100 text-slate-900 text-sm">
                <td colspan="4" class="py-3 px-3 text-right font-black sticky-col-1 bg-slate-100">
                  Total Geral da Proposta (18.999,00 kg):
                </td>
                <td class="py-3 px-2.5 text-right border-l border-slate-200 font-black text-blue-950 bg-blue-100/40 num-tabular" id="tot-br-0">R$ 93.787,27</td>
                <td class="py-3 px-2.5 text-right border-l border-slate-200 font-black text-slate-900 num-tabular" id="tot-br-1">R$ 95.521,22</td>
                <td class="py-3 px-2.5 text-right border-l border-slate-200 font-black text-slate-900 num-tabular" id="tot-br-2">R$ 101.470,10</td>
                <td class="py-3 px-2.5 text-right border-l border-slate-200 font-black text-slate-900 num-tabular" id="tot-br-3">R$ 103.587,31</td>
              </tr>
              <!-- LINHA 4: PRAZOS E CONDIÇÃO COMERCIAL -->
              <tr class="text-[11px] text-slate-600 bg-white">
                <td colspan="4" class="py-2 px-3 text-right font-semibold sticky-col-1 bg-white">Condição de Pagamento / Frete:</td>
                <td class="py-2 px-2.5 text-center border-l border-slate-200 text-blue-900 font-bold">Faturado Usina &bull; CIF</td>
                <td class="py-2 px-2.5 text-center border-l border-slate-200 font-semibold">Antecipado (B014) &bull; CIF</td>
                <td class="py-2 px-2.5 text-center border-l border-slate-200 font-semibold">Faturado &bull; CIF</td>
                <td class="py-2 px-2.5 text-center border-l border-slate-200 font-semibold">Faturado &bull; CIF</td>
              </tr>
              <!-- LINHA 5: STATUS E CLASSIFICAÇÃO -->
              <tr class="text-xs bg-[#0a2540] text-white">
                <td colspan="4" class="py-2.5 px-3 text-right font-bold sticky-col-1 bg-[#0a2540]">Classificação em Barra Reta:</td>
                <td class="py-2.5 px-2.5 text-center border-l border-slate-700" id="stat-br-0"><span class="text-amber-300 font-black">1º Líder Barra (100%)</span></td>
                <td class="py-2.5 px-2.5 text-center border-l border-slate-700" id="stat-br-1"><span>2º Lugar Barra (+1,8%)</span></td>
                <td class="py-2.5 px-2.5 text-center border-l border-slate-700" id="stat-br-2"><span>3º Lugar Barra (+8,2%)</span></td>
                <td class="py-2.5 px-2.5 text-center border-l border-slate-700" id="stat-br-3"><span>4º Lugar Barra (+10,5%)</span></td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      <!-- BOTTOM DASHBOARD SECTION BARRA (3 CARDS LADO A LADO) -->
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <!-- CARD 1: DONUT CHART BARRA -->
        <div class="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div>
            <div class="flex items-center justify-between border-b border-slate-100 pb-2">
              <h4 class="font-bold text-xs text-[#0a2540] uppercase tracking-wider">Distribuição do Custo (Barra)</h4>
              <span class="text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded font-bold">Barra Reta 12m</span>
            </div>
            <div class="flex items-center justify-center my-3">
              <div class="relative w-36 h-36 flex items-center justify-center">
                <svg id="donut-svg-br" class="w-full h-full transform -rotate-90" viewBox="0 0 100 100">
                  <circle cx="50" cy="50" r="38" stroke="#f1f5f9" stroke-width="14" fill="transparent" />
                </svg>
                <div class="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                  <span class="text-xs font-black text-slate-900 num-tabular" id="donut-center-val-br">R$ 94 mil</span>
                  <span class="text-[9px] text-slate-400 font-semibold">Total Barra</span>
                </div>
              </div>
            </div>
          </div>
          <div id="donut-legend-br" class="space-y-1.5 text-xs pt-1 border-t border-slate-100">
            <!-- Gerado via JS -->
          </div>
        </div>

        <!-- CARD 2: COMPOSIÇÃO DE BITOLAS & INSPEÇÃO BARRA -->
        <div class="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div>
            <div class="flex items-center justify-between border-b border-slate-100 pb-2">
              <h4 class="font-bold text-xs text-[#0a2540] uppercase tracking-wider">Composição de Bitolas &amp; Item</h4>
              <span class="text-[10px] text-slate-500 font-semibold">18.999,00 kg total</span>
            </div>
            
            <div class="mt-3">
              <div class="text-[11px] font-semibold text-slate-600 mb-1 flex justify-between">
                <span>Distribuição do Peso por Bitola</span>
                <span class="text-slate-400">12,5mm = 66,7%</span>
              </div>
              <div class="h-3 w-full bg-slate-100 rounded-full overflow-hidden flex" id="composition-bar-br">
                <!-- Gerado via JS -->
              </div>
              <div class="flex justify-between text-[9px] text-slate-400 mt-1" id="composition-legend-br">
                <!-- Gerado via JS -->
              </div>
            </div>

            <!-- DETALHE TÉCNICO DO ITEM INSPECIONADO -->
            <div class="mt-3.5 p-3 bg-slate-50 rounded-xl border border-slate-200/80" id="detail-card-br">
              <div class="flex items-center justify-between mb-1.5">
                <span class="text-[10px] font-bold text-slate-500 uppercase" id="detail-item-cod-br">01.002.0012</span>
                <span class="text-[10px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-800 font-bold" id="detail-item-cat-br">CA-50</span>
              </div>
              <div class="font-bold text-xs text-slate-900" id="detail-item-name-br">FERRO CA-50 12,5MM (1/2")</div>
              <div class="text-[11px] text-slate-500 mt-0.5" id="detail-item-desc-br">Armadura principal de vigas e pilares mestres</div>
              
              <div class="grid grid-cols-3 gap-2 mt-2.5 pt-2 border-t border-slate-200 text-center text-xs">
                <div>
                  <div class="text-[9px] text-slate-400">Volume</div>
                  <div class="font-bold text-slate-800" id="detail-item-vol-br">12.672,00 kg</div>
                </div>
                <div>
                  <div class="text-[9px] text-slate-400">Preço Sel.</div>
                  <div class="font-bold text-blue-900" id="detail-item-price-br">R$ 4,8041/kg</div>
                </div>
                <div>
                  <div class="text-[9px] text-slate-400">Total Item</div>
                  <div class="font-bold text-slate-900" id="detail-item-total-br">R$ 60.876,92</div>
                </div>
              </div>
            </div>
          </div>
          <div class="text-[10px] text-slate-400 text-center mt-2">Clique em qualquer linha da tabela para inspecionar a bitola</div>
        </div>

        <!-- CARD 3: RESUMO DO CENÁRIO BARRA RETA -->
        <div class="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div>
            <div class="flex items-center justify-between border-b border-slate-100 pb-2">
              <h4 class="font-bold text-xs text-[#0a2540] uppercase tracking-wider">Resumo do Cenário (Barra)</h4>
              <span class="text-[10px] bg-blue-100 text-blue-800 px-2 py-0.5 rounded font-bold">100% Barra Reta (12m)</span>
            </div>

            <div class="space-y-2.5 mt-3 text-xs">
              <div class="flex justify-between items-center py-1 border-b border-slate-100">
                <span class="text-slate-500">Custo Total Barra Selecionado:</span>
                <span class="font-black text-slate-900 num-tabular text-sm" id="summary-custo-total-br">R$ 93.787,27</span>
              </div>
              <div class="flex justify-between items-center py-1 border-b border-slate-100">
                <span class="text-slate-500">Menor Mix Item a Item (Barra):</span>
                <span class="font-bold text-blue-800 num-tabular" id="summary-custo-menor-br">R$ 93.783,27</span>
              </div>
              <div class="flex justify-between items-center py-1 border-b border-slate-100">
                <span class="text-slate-500">Economia vs. NORTEFER Barra:</span>
                <span class="font-bold text-emerald-700 num-tabular" id="summary-economia-br">R$ 9.800,04 (-9,5%)</span>
              </div>
              <div class="flex justify-between items-center py-1 border-b border-slate-100">
                <span class="text-slate-500">Atenção Técnica de Canteiro:</span>
                <span class="font-bold text-rose-700">+6,5% perda sucata + M.O. bancada</span>
              </div>
              <div class="flex justify-between items-center py-1 border-b border-slate-100">
                <span class="text-slate-500">Custo Efetivo Real no Canteiro:</span>
                <span class="font-black text-slate-900">R$ 110.834,27</span>
              </div>
              <div class="flex justify-between items-center py-1">
                <span class="text-slate-500">Condição Comercial:</span>
                <span class="font-bold text-blue-900 bg-blue-50 px-2 py-0.5 rounded">Faturado Direto Usina</span>
              </div>
            </div>
          </div>

          <button onclick="switchTab('equalizacao')" class="w-full mt-3 py-2 bg-slate-50 hover:bg-slate-100 text-[#0a2540] border border-slate-200 rounded-xl text-xs font-bold transition flex items-center justify-center space-x-1.5">
            <span>Ver Comparativo Real (CD vs. Barra)</span>
            <span>&rarr;</span>
          </button>
        </div>
      </div>

    </div>

    <!-- ==================== ABA 3: RESUMO EXECUTIVO & RANKING GERAL ==================== -->
    <div id="tab-content-resumo" class="space-y-5 hidden">
      <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-4">
        <div class="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
          <div>
            <h3 class="text-base font-bold text-slate-900">Ranking Geral Consolidado &mdash; Obra 2002 Giardini</h3>
            <p class="text-xs text-slate-500">Comparação executiva das 9 propostas comerciais atualizadas nas duas modalidades (5 Corte e Dobra + 4 Barra Reta)</p>
          </div>
          <span class="px-3 py-1 bg-slate-100 text-slate-700 font-bold text-xs rounded-xl self-start sm:self-auto">
            Volume Total: 18.999,00 kg (19,00 t)
          </span>
        </div>

        <div class="overflow-x-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left text-xs border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-slate-700 font-bold">
                <th class="py-3 px-3 text-center w-16">Posição</th>
                <th class="py-3 px-3">Fornecedor / Usina</th>
                <th class="py-3 px-3 text-center">Modalidade de Fornecimento</th>
                <th class="py-3 px-3 text-right">Valor Cotado Nominal (R$)</th>
                <th class="py-3 px-3 text-right">Diferença vs. Líder Nominal</th>
                <th class="py-3 px-3 text-right">Custo Real Estimado no Canteiro</th>
                <th class="py-3 px-3 text-center">Condição / Escopo</th>
                <th class="py-3 px-3">Parecer de Suprimentos</th>
              </tr>
            </thead>
            <tbody id="executive-ranking-body" class="divide-y divide-slate-100 text-slate-800">
              <!-- Gerado via JS -->
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ==================== ABA 4: SIMULAÇÕES & NEGOCIAÇÃO ==================== -->
    <div id="tab-content-simulacoes" class="space-y-5 hidden">
      <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-5">
        <div class="border-b border-slate-100 pb-3">
          <h3 class="text-base font-bold text-slate-900">Simulador de Contraproposta &amp; Pacotes Comerciais</h3>
          <p class="text-xs text-slate-500">Calcule descontos lineares ou metas globais em qualquer uma das 9 propostas e gere a lista de bitolas para negociação imediata</p>
        </div>

        <!-- CONTROLES DA SIMULAÇÃO -->
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 p-4 bg-slate-50 rounded-2xl border border-slate-200">
          <div>
            <label class="block text-xs font-bold text-slate-700 mb-1">1. Selecione o Fornecedor para Negociar</label>
            <select id="sim-supp-select" onchange="runSim()" class="w-full py-2 px-3 text-xs bg-white border border-slate-300 rounded-xl font-bold text-slate-800 focus:outline-none">
              <optgroup label="Corte e Dobra Industrial (5 Fornecedores)">
                <option value="cd-0" selected>Paranafer CD (1º Aço CD S/ Arame - R$ 101.447,64)</option>
                <option value="cd-1">ArcelorMittal CD (1º Completo - R$ 109.038,45)</option>
                <option value="cd-2">Gerdau CD (2º Completo - R$ 110.756,44)</option>
                <option value="cd-4">NORTEFER CD (3º Completo - R$ 112.381,55)</option>
                <option value="cd-3">Kauefer CD (4º Completo - R$ 112.884,65)</option>
              </optgroup>
              <optgroup label="Barra Reta 12m (4 Fornecedores)">
                <option value="br-0">ArcelorMittal Barra (1º Líder Barra - R$ 93.787,27)</option>
                <option value="br-1">Gerdau Barra (2º Barra Completo - R$ 95.521,22)</option>
                <option value="br-2">Kauefer Barra (3º Barra Completo - R$ 101.470,10)</option>
                <option value="br-3">NORTEFER Barra (4º Barra Completo - R$ 103.587,31)</option>
              </optgroup>
            </select>
          </div>

          <div>
            <label class="block text-xs font-bold text-slate-700 mb-1">2. Percentual de Desconto Solicitado (%)</label>
            <div class="flex items-center space-x-2">
              <input type="number" id="sim-disc-input" value="3.0" min="0" max="25" step="0.5" oninput="runSimFromInput(this.value)" class="w-24 py-2 px-3 text-xs bg-white border border-slate-300 rounded-xl font-bold text-slate-800 text-center focus:outline-none">
              <input type="range" id="sim-disc-slider" value="3.0" min="0" max="15" step="0.5" oninput="runSimFromSlider(this.value)" class="flex-1 accent-[#0a2540]">
            </div>
            <!-- BARRAS DE SIMULAÇÃO RÁPIDA -->
            <div class="flex space-x-1.5 mt-2">
              <button onclick="setFastSim(2.0)" class="px-2 py-0.5 bg-white hover:bg-slate-200 border border-slate-300 rounded text-[11px] font-bold text-slate-700 transition">-2%</button>
              <button onclick="setFastSim(3.0)" class="px-2 py-0.5 bg-white hover:bg-slate-200 border border-slate-300 rounded text-[11px] font-bold text-slate-700 transition">-3%</button>
              <button onclick="setFastSim(5.0)" class="px-2 py-0.5 bg-white hover:bg-slate-200 border border-slate-300 rounded text-[11px] font-bold text-slate-700 transition">-5%</button>
              <button onclick="setFastSim(7.0)" class="px-2 py-0.5 bg-white hover:bg-slate-200 border border-slate-300 rounded text-[11px] font-bold text-slate-700 transition">-7%</button>
              <button onclick="setFastSim(10.0)" class="px-2 py-0.5 bg-white hover:bg-slate-200 border border-slate-300 rounded text-[11px] font-bold text-slate-700 transition">-10%</button>
            </div>
          </div>

          <!-- RESUMO NUMÉRICO DA SIMULAÇÃO -->
          <div class="bg-white p-3 rounded-xl border border-slate-200 flex flex-col justify-between">
            <div class="flex justify-between items-center text-xs">
              <span class="text-slate-500">Valor Cotado Original:</span>
              <span class="font-bold text-slate-700 num-tabular" id="sim-val-original">R$ 101.447,64</span>
            </div>
            <div class="flex justify-between items-center text-xs">
              <span class="text-emerald-700 font-semibold">Desconto Negociado:</span>
              <span class="font-bold text-emerald-700 num-tabular" id="sim-val-desconto">-R$ 3.043,43</span>
            </div>
            <div class="flex justify-between items-center text-xs pt-1 border-t border-slate-100">
              <span class="font-extrabold text-[#0a2540]">Valor Alvo da Contraproposta:</span>
              <span class="font-black text-slate-900 num-tabular text-sm text-[#0a2540]" id="sim-val-final">R$ 98.404,21</span>
            </div>
          </div>
        </div>

        <!-- TABELA DETALHADA DE BITOLAS COM O DESCONTO SIMULADO -->
        <div class="space-y-3">
          <div class="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2">
            <div>
              <h4 class="font-bold text-xs text-[#0a2540] uppercase tracking-wider">Tabela de Preços Unitários com Desconto Simulado</h4>
              <p class="text-[11px] text-slate-500">Valores por quilo (R$/kg) resultantes para envio formal ao fornecedor</p>
            </div>
            <button onclick="copyNegotiatedProposal()" id="btn-copy-sim" class="px-4 py-2 bg-[#0a2540] hover:bg-[#12365c] text-white rounded-xl text-xs font-bold transition flex items-center justify-center space-x-1.5 shadow-xs">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 17.25v3.375c0 .621-.504 1.125-1.125 1.125h-9.75a1.125 1.125 0 01-1.125-1.125V7.875c0-.621.504-1.125 1.125-1.125H6.75a9.06 9.06 0 011.5.124m7.5 10.376h3.375c.621 0 1.125-.504 1.125-1.125V11.25c0-4.46-3.243-8.161-7.5-8.876a9.06 9.06 0 00-1.5-.124H9.375c-.621 0-1.125.504-1.125 1.125v3.5m7.5 10.375H9.375a1.125 1.125 0 01-1.125-1.125v-9.25m12 6.625v-1.875a3.375 3.375 0 00-3.375-3.375h-1.5a1.125 1.125 0 01-1.125-1.125v-1.5a3.375 3.375 0 00-3.375-3.375H9.75"/>
              </svg>
              <span>Copiar Contraproposta</span>
            </button>
          </div>

          <div class="overflow-x-auto border border-slate-200 rounded-xl">
            <table class="w-full text-left text-xs border-collapse">
              <thead>
                <tr class="bg-slate-50 border-b border-slate-200 text-slate-700 font-bold">
                  <th class="py-2.5 px-3">Código</th>
                  <th class="py-2.5 px-3">Descrição da Bitola</th>
                  <th class="py-2.5 px-2 text-center">Unid.</th>
                  <th class="py-2.5 px-3 text-right">Volume (kg)</th>
                  <th class="py-2.5 px-3 text-right text-slate-500">Preço Cotado</th>
                  <th class="py-2.5 px-3 text-right bg-emerald-50 text-emerald-950 font-black">Preço Negociado</th>
                  <th class="py-2.5 px-3 text-right text-slate-500">Total Cotado</th>
                  <th class="py-2.5 px-3 text-right bg-emerald-50 text-emerald-950 font-black">Total Negociado</th>
                  <th class="py-2.5 px-3 text-right text-emerald-700 font-bold">Economia no Item</th>
                </tr>
              </thead>
              <tbody id="sim-table-body" class="divide-y divide-slate-100 text-slate-800">
                <!-- Gerado via JS -->
              </tbody>
              <tfoot class="bg-slate-50 font-bold border-t-2 border-slate-200">
                <tr class="text-slate-900">
                  <td colspan="4" class="py-3 px-3 text-right font-black">Totais Consolidados da Contraproposta:</td>
                  <td class="py-3 px-3 text-right num-tabular" id="sim-foot-avg-orig">-</td>
                  <td class="py-3 px-3 text-right bg-emerald-100/50 text-emerald-900 num-tabular" id="sim-foot-avg-neg">-</td>
                  <td class="py-3 px-3 text-right num-tabular" id="sim-foot-tot-orig">R$ 101.447,64</td>
                  <td class="py-3 px-3 text-right bg-emerald-100/50 text-emerald-950 font-black num-tabular" id="sim-foot-tot-neg">R$ 98.404,21</td>
                  <td class="py-3 px-3 text-right text-emerald-700 font-black num-tabular" id="sim-foot-tot-eco">-R$ 3.043,43</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      </div>
    </div>

    <!-- ==================== ABA 5: EQUALIZAÇÃO TÉCNICA (CD VS. BARRA RETA) ==================== -->
    <div id="tab-content-equalizacao" class="space-y-5 hidden">
      <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-4">
        <div>
          <h3 class="text-base font-bold text-slate-900">Equalização Técnica: Corte e Dobra Industrial vs. Barra Reta</h3>
          <p class="text-xs text-slate-600 leading-relaxed">
            Demonstração matemática dos impactos reais no canteiro da Obra Giardini (perdas de ponta, mão de obra de bancada, logística e segurança):
          </p>
        </div>

        <div class="overflow-x-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left text-xs border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-slate-700 font-bold">
                <th class="py-3 px-3 w-1/4">Parâmetro de Engenharia</th>
                <th class="py-3 px-3 text-center">Barra Reta (12m)</th>
                <th class="py-3 px-3 text-center bg-cyan-50/50">Corte e Dobra Industrial</th>
                <th class="py-3 px-3">Impacto Prático no Canteiro Giardini (Av. Londrina, 494)</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-100">
              <tr>
                <td class="py-2.5 px-3 font-semibold text-slate-800">Perda de Aço / Sucata de Pontas</td>
                <td class="py-2.5 px-3 text-center text-rose-700 font-bold">5,0% a 8,0% (1.209 kg = +R$ 5.888)</td>
                <td class="py-2.5 px-3 text-center text-emerald-700 font-bold bg-cyan-50/30">0,0% (Zero Perda de Ponta)</td>
                <td class="py-2.5 px-3 text-slate-600">No corte manual, as pontas das barras de 12m são descartadas como sucata. No CD, a construtora paga pelo peso exato de projeto.</td>
              </tr>
              <tr>
                <td class="py-2.5 px-3 font-semibold text-slate-800">Mão de Obra de Armadores no Canteiro</td>
                <td class="py-2.5 px-3 text-center text-rose-700 font-bold">Equipe Dedicada (~R$ 11.159)</td>
                <td class="py-2.5 px-3 text-center text-emerald-700 font-bold bg-cyan-50/30">R$ 0,00 (Processado em Fábrica)</td>
                <td class="py-2.5 px-3 text-slate-600">Elimina a bancada artesanal de armação. Os armadores da obra focam 100% do tempo no posicionamento e amarração nas formas.</td>
              </tr>
              <tr>
                <td class="py-2.5 px-3 font-semibold text-slate-800">Espaço Físico Ocupado no Canteiro</td>
                <td class="py-2.5 px-3 text-center text-rose-700 font-semibold">&gt; 60 m² de Canteiro Ocupado</td>
                <td class="py-2.5 px-3 text-center text-emerald-700 font-bold bg-cyan-50/30">Just-in-Time (Canteiro Livre)</td>
                <td class="py-2.5 px-3 text-slate-600">Barras de 12m exigem área coberta de 14m de comprimento. O CD é entregue paletizado por pavimento/etapa pronta.</td>
              </tr>
              <tr>
                <td class="py-2.5 px-3 font-semibold text-slate-800">Produtividade &amp; Ciclo de Formas</td>
                <td class="py-2.5 px-3 text-center text-slate-600">Gargalo do Cronograma</td>
                <td class="py-2.5 px-3 text-center text-emerald-700 font-bold bg-cyan-50/30">+35% de Agilidade na Montagem</td>
                <td class="py-2.5 px-3 text-slate-600">Peças chegam identificadas por elemento estrutural (viga/pilar). Reduz drasticamente o tempo de liberação da forma para concretagem.</td>
              </tr>
              <tr>
                <td class="py-2.5 px-3 font-semibold text-slate-800">Segurança do Trabalho (NR-18)</td>
                <td class="py-2.5 px-3 text-center text-rose-700 font-semibold">Riscos com Policorte e Ruído</td>
                <td class="py-2.5 px-3 text-center text-emerald-700 font-bold bg-cyan-50/30">Ambiente Limpo e Seguro</td>
                <td class="py-2.5 px-3 text-slate-600">Elimina policortes de alta rotação, centelhas e esforço ergonômico repetitivo de dobra manual de bitolas grossas (12.5mm e 16mm).</td>
              </tr>
              <tr>
                <td class="py-2.5 px-3 font-semibold text-slate-800">Rastreabilidade e Precisão Dimensional</td>
                <td class="py-2.5 px-3 text-center text-slate-600">Variação Manual Humana</td>
                <td class="py-2.5 px-3 text-center text-emerald-700 font-bold bg-cyan-50/30">Etiqueta por Elemento (V101, P12)</td>
                <td class="py-2.5 px-3 text-slate-600">Cada feixe possui etiqueta industrial com prancha, elemento e bitola. Dobras a frio em máquinas CNC conforme NBR 6118.</td>
              </tr>
            </tbody>
          </table>
        </div>

        <!-- BALANÇO FINANCEIRO REAL NO CANTEIRO -->
        <div class="p-4 bg-slate-50 rounded-2xl border border-slate-200 space-y-2">
          <h4 class="font-bold text-xs text-[#0a2540] uppercase tracking-wider">Balanço Comparativo Final: Custo Real do Aço no Canteiro</h4>
          <div class="grid grid-cols-1 md:grid-cols-3 gap-3 pt-1 text-xs">
            <div class="p-3 bg-white rounded-xl border border-slate-200">
              <div class="text-slate-500 font-semibold">Custo Nominal Barra Reta (Arcelor):</div>
              <div class="text-base font-black text-slate-800 mt-1">R$ 93.787,27</div>
              <div class="text-[11px] text-rose-700 mt-1 font-semibold">+ R$ 5.888 (Sucata 6,5%)<br>+ R$ 11.159 (M.O. Bancada)<br>= <strong>R$ 110.834,27 Real</strong></div>
            </div>

            <div class="p-3 bg-cyan-50/70 rounded-xl border border-cyan-200">
              <div class="text-cyan-900 font-semibold">Custo Corte e Dobra (Paranafer + Arame):</div>
              <div class="text-base font-black text-cyan-950 mt-1">R$ 104.687,64</div>
              <div class="text-[11px] text-cyan-800 mt-1 font-semibold">+ R$ 0,00 Sucata (Zero perda)<br>+ R$ 0,00 M.O. Bancada (Incluso)<br>= <strong>R$ 104.687,64 Real</strong></div>
            </div>

            <div class="p-3 bg-emerald-50 rounded-xl border border-emerald-200 flex flex-col justify-between">
              <div>
                <div class="text-emerald-800 font-bold">Economia Líquida Real para a Evense:</div>
                <div class="text-lg font-black text-emerald-700 mt-1">R$ 6.146,63</div>
              </div>
              <div class="text-[11px] text-emerald-800 font-semibold">
                Mesmo na proposta 100% Arcelor CD (R$ 109.038,45), o Corte e Dobra ainda é <strong>R$ 1.795,82 mais barato</strong> que o custo real da Barra Reta.
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- ==================== ABA 6: ANÁLISE TÉCNICA & NORMAS ==================== -->
    <div id="tab-content-tecnica" class="space-y-5 hidden">
      <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-4">
        <div>
          <h3 class="text-base font-bold text-slate-900">Auditoria Técnica de Projetos: Exigência vs. Oferta Comercial</h3>
          <p class="text-xs text-slate-600 leading-relaxed">
            Confronto dos requisitos de projeto estrutural (Zocco Engenharia R01 e NBR 7480/NBR 6118) com as propostas comerciais homologadas:
          </p>
        </div>

        <div class="overflow-x-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left text-xs border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-slate-700 font-bold">
                <th class="py-3 px-3 w-1/4">Elemento Estrutural / Bitola</th>
                <th class="py-3 px-3 w-1/4">Exigência do Projeto &amp; Norma</th>
                <th class="py-3 px-3 w-1/4">Oferta Cotada &amp; Melhor Preço</th>
                <th class="py-3 px-2 text-center w-28">Status Técnico</th>
                <th class="py-3 px-3 w-1/4">Parecer da Engenharia</th>
              </tr>
            </thead>
            <tbody id="audit-table-body" class="divide-y divide-slate-100">
              <!-- Gerado via JavaScript -->
            </tbody>
          </table>
        </div>
      </div>

      <!-- REFERÊNCIA DIMENSIONAL E TOLERÂNCIAS NBR 7480 -->
      <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-4">
        <h3 class="text-base font-bold text-slate-900">Massa Linear Teórica e Tolerâncias ABNT NBR 7480</h3>
        <p class="text-xs text-slate-500">Parâmetros normativos para controle de recebimento por balança e verificação de massa linear (kg/m):</p>
        
        <div class="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-3 text-xs">
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200 text-center">
            <div class="text-[10px] text-slate-400 font-bold uppercase">CA-60 5,0mm</div>
            <div class="text-xs font-black text-[#0a2540] mt-0.5">0,154 kg/m</div>
            <div class="text-[10px] text-slate-500 mt-0.5">&plusmn; 6,0%</div>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200 text-center">
            <div class="text-[10px] text-slate-400 font-bold uppercase">CA-50 6,3mm</div>
            <div class="text-xs font-black text-[#0a2540] mt-0.5">0,245 kg/m</div>
            <div class="text-[10px] text-slate-500 mt-0.5">&plusmn; 6,0%</div>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200 text-center">
            <div class="text-[10px] text-slate-400 font-bold uppercase">CA-50 8,0mm</div>
            <div class="text-xs font-black text-[#0a2540] mt-0.5">0,395 kg/m</div>
            <div class="text-[10px] text-slate-500 mt-0.5">&plusmn; 5,0%</div>
          </div>
          <div class="p-3 bg-blue-50/70 rounded-xl border border-blue-200 text-center">
            <div class="text-[10px] text-blue-900 font-bold uppercase">CA-50 10,0mm</div>
            <div class="text-xs font-black text-blue-950 mt-0.5">0,617 kg/m</div>
            <div class="text-[10px] text-blue-700 mt-0.5">&plusmn; 5,0%</div>
          </div>
          <div class="p-3 bg-blue-100/70 rounded-xl border border-blue-300 text-center">
            <div class="text-[10px] text-blue-950 font-black uppercase">CA-50 12,5mm</div>
            <div class="text-xs font-black text-blue-950 mt-0.5">0,963 kg/m</div>
            <div class="text-[10px] text-blue-800 mt-0.5 font-bold">&plusmn; 4,0%</div>
          </div>
          <div class="p-3 bg-blue-50/70 rounded-xl border border-blue-200 text-center">
            <div class="text-[10px] text-blue-900 font-bold uppercase">CA-50 16,0mm</div>
            <div class="text-xs font-black text-blue-950 mt-0.5">1,578 kg/m</div>
            <div class="text-[10px] text-blue-700 mt-0.5">&plusmn; 4,0%</div>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200 text-center">
            <div class="text-[10px] text-slate-400 font-bold uppercase">CA-50 20,0mm</div>
            <div class="text-xs font-black text-[#0a2540] mt-0.5">2,466 kg/m</div>
            <div class="text-[10px] text-slate-500 mt-0.5">&plusmn; 4,0%</div>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200 text-center">
            <div class="text-[10px] text-slate-400 font-bold uppercase">CA-50 25,0mm</div>
            <div class="text-xs font-black text-[#0a2540] mt-0.5">3,853 kg/m</div>
            <div class="text-[10px] text-slate-500 mt-0.5">&plusmn; 4,0%</div>
          </div>
        </div>
      </div>
    </div>

    <!-- ==================== ABA 7: IMPORTAÇÃO SIENGE ERP ==================== -->
    <div id="tab-content-sienge" class="space-y-5 hidden">
      <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-4">
        <div class="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
          <div>
            <h3 class="text-base font-bold text-slate-900">Integração com ERP Sienge &mdash; Módulo de Compras</h3>
            <p class="text-xs text-slate-500">Tabela formatada com códigos oficiais Sienge e valores unitários homologados na seleção atual</p>
          </div>
          <div class="flex items-center space-x-2">
            <select id="sienge-modality-select" onchange="renderSiengeTable()" class="py-2 px-3 text-xs bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-800 focus:outline-none">
              <option value="cd" selected>Exportar Seleção Corte e Dobra</option>
              <option value="br">Exportar Seleção Barra Reta</option>
            </select>
            <button onclick="copySiengeData()" id="btn-copy-sienge" class="px-4 py-2 bg-[#0a2540] hover:bg-[#12365c] text-white rounded-xl text-xs font-bold transition flex items-center space-x-1.5 shadow-xs">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 17.25v3.375c0 .621-.504 1.125-1.125 1.125h-9.75a1.125 1.125 0 01-1.125-1.125V7.875c0-.621.504-1.125 1.125-1.125H6.75a9.06 9.06 0 011.5.124m7.5 10.376h3.375c.621 0 1.125-.504 1.125-1.125V11.25c0-4.46-3.243-8.161-7.5-8.876a9.06 9.06 0 00-1.5-.124H9.375c-.621 0-1.125.504-1.125 1.125v3.5m7.5 10.375H9.375a1.125 1.125 0 01-1.125-1.125v-9.25m12 6.625v-1.875a3.375 3.375 0 00-3.375-3.375h-1.5a1.125 1.125 0 01-1.125-1.125v-1.5a3.375 3.375 0 00-3.375-3.375H9.75"/>
              </svg>
              <span>Copiar TSV para Sienge</span>
            </button>
          </div>
        </div>

        <div class="overflow-x-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left text-xs border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-slate-700 font-bold">
                <th class="py-2.5 px-3">Código Sienge</th>
                <th class="py-2.5 px-3">Descrição do Insumo</th>
                <th class="py-2.5 px-2 text-center">Unid.</th>
                <th class="py-2.5 px-3 text-right">Volume (kg)</th>
                <th class="py-2.5 px-3">Fornecedor Homologado</th>
                <th class="py-2.5 px-3 text-right">Preço Unit. (R$)</th>
                <th class="py-2.5 px-3 text-right">Valor Total (R$)</th>
              </tr>
            </thead>
            <tbody id="sienge-table-body" class="divide-y divide-slate-100">
              <!-- Gerado via JavaScript -->
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ==================== ABA 8: GRÁFICOS COMPARATIVOS ==================== -->
    <div id="tab-content-graficos" class="space-y-5 hidden">
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-5">
        <!-- BARRAS COMPARATIVAS -->
        <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-4">
          <div>
            <h4 class="font-bold text-sm text-[#0a2540]">Custo Total por Fornecedor (R$) &mdash; 9 Propostas</h4>
            <p class="text-xs text-slate-400">Comparativo das 9 cotações (5 Corte e Dobra + 4 Barra Reta) para o pacote da Obra 2002 Giardini</p>
          </div>
          <div id="chart-bars-container" class="space-y-3 pt-2">
            <!-- Gerado via JS -->
          </div>
        </div>

        <!-- COMPOSIÇÃO DE VOLUME -->
        <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-4">
          <div>
            <h4 class="font-bold text-sm text-[#0a2540]">Distribuição Estrutural do Volume (kg)</h4>
            <p class="text-xs text-slate-400">Composição do consumo de aço conforme projeto executivo Zocco R01</p>
          </div>
          <div id="chart-comp-container" class="space-y-3 pt-2">
            <!-- Gerado via JS -->
          </div>
        </div>
      </div>
    </div>

    <!-- 5. RODAPÉ INSTITUCIONAL COM BLOCO DE ASSINATURAS E APROVAÇÃO -->
    <footer class="bg-white rounded-2xl border border-slate-200/80 p-5 shadow-xs space-y-4">
      <div class="flex flex-col sm:flex-row items-center justify-between border-b border-slate-100 pb-3 gap-2">
        <div class="flex items-center space-x-2 text-xs">
          <img src="data:image/png;base64,{LOGO_BASE64}" alt="Evense" class="h-5 w-auto object-contain">
          <span class="font-bold text-slate-700">EVENSE EMPREENDIMENTOS &bull; DEPARTAMENTO DE ENGENHARIA E SUPRIMENTOS</span>
        </div>
        <div class="text-[11px] text-slate-400">
          Obra 2002 Giardini &bull; Normas: ABNT NBR 7480 / NBR 6118 &bull; Atualização: 22/10/2026
        </div>
      </div>

      <div class="grid grid-cols-1 sm:grid-cols-3 gap-6 text-center pt-1">
        <div>
          <div class="border-b border-slate-300 pb-1 mb-1 font-bold text-slate-800 text-xs">Eng. Marcio Gracional</div>
          <div class="text-[10px] text-slate-500 font-medium">Eng. de Planejamento / Suprimentos</div>
        </div>
        <div>
          <div class="border-b border-slate-300 pb-1 mb-1 font-bold text-slate-800 text-xs">Coordenação de Obras</div>
          <div class="text-[10px] text-slate-400">Edifício Residencial Giardini</div>
        </div>
        <div>
          <div class="border-b border-slate-300 pb-1 mb-1 font-bold text-slate-800 text-xs">Diretoria de Engenharia</div>
          <div class="text-[10px] text-slate-400">Evense Empreendimentos</div>
        </div>
      </div>
    </footer>

  </div>

  <!-- JAVASCRIPT ENGINE (MOTOR REATIVO COMPLETO) -->
  <script>
    // DADOS MESTRES DOS ITENS (ATUALIZADO COM NORTEFER REV00 E GERDAU PT2)
    // pricesCD: [Paranafer, ArcelorMittal CD, Gerdau CD, Kauefer CD, NORTEFER CD]
    // pricesBR: [ArcelorMittal Barra, Gerdau Barra, Kauefer Barra, NORTEFER Barra]
    const itemsData = [
      {{
        "id": 1,
        "cod": "01.002.0005",
        "nome": "FERRO CA-60 5,0MM (3/16\\\")",
        "desc": "Estribos de vigas, pilares e telas de amarração",
        "cat": "ca60",
        "unid": "kg",
        "qtd": 1687.0,
        "pricesCD": [5.80, 6.3016, 6.44, 6.15, 6.05],
        "pricesBR": [5.4816, 5.52, 5.90, 6.05]
      }},
      {{
        "id": 2,
        "cod": "01.002.0006",
        "nome": "FERRO CA-50 6,3MM (1/4\\\")",
        "desc": "Armadura secundária e estribos especiais (Ref. Projeto)",
        "cat": "ca50",
        "unid": "kg",
        "qtd": 0.0,
        "pricesCD": [5.75, 5.9530, 6.10, 6.05, 6.05],
        "pricesBR": [5.13335, 5.23, 5.49, 5.44]
      }},
      {{
        "id": 3,
        "cod": "01.002.0008",
        "nome": "FERRO CA-50 8,0MM (5/16\\\")",
        "desc": "Armadura de distribuição e vigamento leve (Ref. Projeto)",
        "cat": "ca50",
        "unid": "kg",
        "qtd": 0.0,
        "pricesCD": [5.62, 5.9530, 6.10, 6.05, 6.05],
        "pricesBR": [5.13335, 5.23, 5.49, 5.46]
      }},
      {{
        "id": 4,
        "cod": "01.002.0010",
        "nome": "FERRO CA-50 10,0MM (3/8\\\")",
        "desc": "Armadura longitudinal de pilares e vigas",
        "cat": "ca50",
        "unid": "kg",
        "qtd": 518.0,
        "pricesCD": [5.42, 5.72003, 5.82, 5.80, 5.85],
        "pricesBR": [4.90003, 5.04, 5.15, 5.33]
      }},
      {{
        "id": 5,
        "cod": "01.002.0012",
        "nome": "FERRO CA-50 12,5MM (1/2\\\")",
        "desc": "Armadura principal de vigas e pilares mestres",
        "cat": "ca50",
        "unid": "kg",
        "qtd": 12672.0,
        "pricesCD": [5.42, 5.62405, 5.65, 5.80, 5.85],
        "pricesBR": [4.80405, 4.84, 5.15, 5.33]
      }},
      {{
        "id": 6,
        "cod": "01.002.0016",
        "nome": "FERRO CA-50 16,0MM (5/8\\\")",
        "desc": "Armadura pesada de fundações, pilares e transições",
        "cat": "ca50",
        "unid": "kg",
        "qtd": 3722.0,
        "pricesCD": [5.42, 5.62405, 5.65, 5.80, 5.85],
        "pricesBR": [4.80405, 4.84, 5.15, 5.33]
      }},
      {{
        "id": 7,
        "cod": "01.002.0020",
        "nome": "FERRO CA-50 20,0MM (3/4\\\")",
        "desc": "Armadura de blocos de coroamento e transições (Ref. Projeto)",
        "cat": "ca50",
        "unid": "kg",
        "qtd": 0.0,
        "pricesCD": [5.60, 5.62405, 5.65, 5.80, 5.85],
        "pricesBR": [4.80405, 4.84, 5.15, 5.23]
      }},
      {{
        "id": 8,
        "cod": "01.002.0025",
        "nome": "FERRO CA-50 25,0MM (1\\\")",
        "desc": "Armadura de alta solicitação e pilares centrais (Ref. Projeto)",
        "cat": "ca50",
        "unid": "kg",
        "qtd": 0.0,
        "pricesCD": [5.60, 5.62405, 5.65, 5.80, 5.85],
        "pricesBR": [4.80405, 4.84, 5.15, 5.23]
      }},
      {{
        "id": 9,
        "cod": "01.005.0016",
        "nome": "ARAME RECOZIDO BWG 16 NF",
        "desc": "Fixação e amarração de armaduras estruturais",
        "cat": "arame",
        "unid": "kg",
        "qtd": 400.0,
        "pricesCD": [null, 8.11, 10.62825, 11.05, 8.10],
        "pricesBR": [8.11, 10.62825, 11.05, 8.10]
      }}
    ];

    // DEFINIÇÃO DOS 5 FORNECEDORES DE CORTE E DOBRA
    const suppliersCD = [
      {{ "nome": "Paranafer", "modalidade": "Corte e Dobra", "prazo": "30 Dias Boleto", "color": "#0891b2" }},
      {{ "nome": "ArcelorMittal (CD)", "modalidade": "Corte e Dobra", "prazo": "15 Dias Corridos", "color": "#0284c7" }},
      {{ "nome": "Gerdau (CD)", "modalidade": "Corte e Dobra", "prazo": "Antecipado (B014)", "color": "#2563eb" }},
      {{ "nome": "Kauefer (CD)", "modalidade": "Corte e Dobra", "prazo": "Faturado Distribuidor", "color": "#64748b" }},
      {{ "nome": "NORTEFER (CD)", "modalidade": "Corte e Dobra", "prazo": "Faturado Programado", "color": "#475569" }}
    ];

    // DEFINIÇÃO DOS 4 FORNECEDORES DE BARRA RETA
    const suppliersBR = [
      {{ "nome": "ArcelorMittal", "modalidade": "Barra Reta (12m)", "prazo": "Faturado Usina", "color": "#1d4ed8" }},
      {{ "nome": "Gerdau", "modalidade": "Barra Reta (12m)", "prazo": "Antecipado (B014)", "color": "#2563eb" }},
      {{ "nome": "Kauefer", "modalidade": "Barra Reta (12m)", "prazo": "Faturado Distribuidor", "color": "#64748b" }},
      {{ "nome": "NORTEFER", "modalidade": "Barra Reta (12m)", "prazo": "Faturado Programado", "color": "#0891b2" }}
    ];

    // ESTADO DE SELEÇÃO INICIAL: Menor preço item a item em CD (Paranafer no aço + Nortefer no arame) e Arcelor em Barra
    let selectedCD = {{ 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 4 }};
    let selectedBR = {{ 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0 }};

    let activeTab = "corte-dobra";
    let inspectedItemId = 5; // CA-50 12,5mm por padrão

    let searchCD = "";
    let filterCatCD = "all";
    let onlyActiveCD = false;

    let searchBR = "";
    let filterCatBR = "all";
    let onlyActiveBR = false;

    // FORMATAÇÃO BRL
    function formatBRL(val) {{
      if (val === null || val === undefined || isNaN(val)) return "-";
      return val.toLocaleString("pt-BR", {{ style: "currency", currency: "BRL" }});
    }}

    function formatNum(val, dec = 2) {{
      if (val === null || val === undefined || isNaN(val)) return "0,00";
      return val.toLocaleString("pt-BR", {{ minimumFractionDigits: dec, maximumFractionDigits: dec }});
    }}

    // NAVEGAÇÃO DE ABAS
    function switchTab(tabId) {{
      activeTab = tabId;
      const tabs = ["corte-dobra", "barra-reta", "resumo", "simulacoes", "equalizacao", "tecnica", "sienge", "graficos"];
      tabs.forEach(t => {{
        const content = document.getElementById(`tab-content-${{t}}`);
        const btn = document.getElementById(`tab-btn-${{t}}`);
        if (content) content.classList.add("hidden");
        if (btn) btn.classList.remove("active");
      }});
      const targetContent = document.getElementById(`tab-content-${{tabId}}`);
      const targetBtn = document.getElementById(`tab-btn-${{tabId}}`);
      if (targetContent) targetContent.classList.remove("hidden");
      if (targetBtn) targetBtn.classList.add("active");

      updateTopKpis();
    }}

    // ATUALIZAÇÃO DOS TOP METRIC CARDS
    function updateTopKpis() {{
      if (activeTab === "barra-reta") {{
        calculateTotalsBR();
      }} else {{
        calculateTotalsCD();
      }}
    }}

    // ==================== LÓGICA DE CORTE E DOBRA ====================
    function handleSearchCD(val) {{
      searchCD = val.toLowerCase().trim();
      renderTableCD();
    }}
    function setFilterCD(cat) {{
      filterCatCD = cat;
      renderTableCD();
    }}
    function toggleOnlyActiveCD(checked) {{
      onlyActiveCD = checked;
      renderTableCD();
    }}
    function selectSingleSupplierCD(sIdx) {{
      itemsData.forEach(item => {{
        if (item.pricesCD[sIdx] !== null) {{
          selectedCD[item.id] = sIdx;
        }}
      }});
      renderTableCD();
    }}
    function selectBestPricesAllCD() {{
      itemsData.forEach(item => {{
        let bestIdx = -1;
        let minP = Infinity;
        item.pricesCD.forEach((p, idx) => {{
          if (p !== null && p < minP) {{
            minP = p;
            bestIdx = idx;
          }}
        }});
        if (bestIdx !== -1) selectedCD[item.id] = bestIdx;
      }});
      renderTableCD();
    }}
    function handleRadioSelectCD(itemId, sIdx) {{
      selectedCD[itemId] = sIdx;
      inspectedItemId = itemId;
      renderTableCD();
    }}

    function renderTableCD() {{
      const tbody = document.getElementById("table-body-cd");
      if (!tbody) return;
      tbody.innerHTML = "";

      const filtered = itemsData.filter(item => {{
        if (onlyActiveCD && item.qtd <= 0) return false;
        if (filterCatCD !== "all" && item.cat !== filterCatCD) return false;
        if (searchCD) {{
          const term = searchCD;
          const matchCod = item.cod.toLowerCase().includes(term);
          const matchName = item.nome.toLowerCase().includes(term);
          const matchDesc = item.desc.toLowerCase().includes(term);
          if (!matchCod && !matchName && !matchDesc) return false;
        }}
        return true;
      }});

      filtered.forEach(item => {{
        const isInspected = (item.id === inspectedItemId);
        let minP = Infinity;
        item.pricesCD.forEach(p => {{
          if (p !== null && p < minP) minP = p;
        }});

        const tr = document.createElement("tr");
        tr.className = `hover:bg-slate-50 transition cursor-pointer ${{isInspected ? "bg-cyan-50/20" : ""}}`;
        tr.onclick = () => {{
          inspectedItemId = item.id;
          updateDetailCardCD();
          updateDetailCardBR();
        }};

        let html = `
          <td class="py-2.5 px-2.5 font-semibold text-slate-500 font-mono text-[11px] sticky-col-1 bg-white">${{item.cod}}</td>
          <td class="py-2.5 px-2.5 sticky-col-2 bg-white">
            <div class="font-bold text-slate-900">${{item.nome}}</div>
            <div class="text-[10px] text-slate-400 truncate max-w-[200px]">${{item.desc}}</div>
          </td>
          <td class="py-2.5 px-1.5 text-center font-bold text-slate-500">${{item.unid}}</td>
          <td class="py-2.5 px-2.5 text-right">
            <input type="number" step="0.01" value="${{item.qtd.toFixed(2)}}" onchange="updateItemQty(${{item.id}}, this.value)" class="w-20 py-1 px-1.5 text-right font-bold text-slate-800 bg-slate-50 hover:bg-slate-100 focus:bg-white border border-slate-200 rounded-lg text-xs num-tabular focus:outline-none focus:ring-1 focus:ring-[#0a2540]">
          </td>
        `;

        // 5 Colunas Fornecedores CD
        suppliersCD.forEach((_, sIdx) => {{
          const p = item.pricesCD[sIdx];
          const isSelected = (selectedCD[item.id] === sIdx);
          const isBest = (p !== null && p === minP);
          const tot = (p !== null && item.qtd > 0) ? (p * item.qtd) : 0;

          let diffBadge = "";
          if (p !== null && minP < Infinity && p > minP) {{
            const dVal = p - minP;
            const dPct = (dVal / minP) * 100;
            diffBadge = `<span class="diff-pill diff-pill-higher">+${{dPct.toFixed(1)}}%</span>`;
          }} else if (isBest) {{
            diffBadge = `<span class="diff-pill diff-pill-best">Menor</span>`;
          }}

          if (p === null) {{
            html += `
              <td class="py-2.5 px-2.5 text-center border-l border-slate-100 bg-slate-50/40 text-slate-400 italic text-[11px]">
                Não cotado
              </td>
            `;
          }} else {{
            html += `
              <td class="py-2.5 px-2.5 border-l border-slate-100 ${{isSelected ? "cell-selected" : ""}} ${{isBest ? "cell-best" : ""}}" onclick="handleRadioSelectCD(${{item.id}}, ${{sIdx}})">
                <div class="flex items-center justify-between">
                  <div class="saas-radio ${{isSelected ? "checked" : ""}}"></div>
                  <div class="text-right flex-1 ml-1.5">
                    <div class="flex items-center justify-end space-x-1">
                      <span class="font-extrabold text-slate-900 num-tabular text-xs">${{formatBRL(p)}}</span>
                      ${{diffBadge}}
                    </div>
                    <div class="text-[10px] text-slate-500 num-tabular mt-0.5">
                      ${{item.qtd > 0 ? formatBRL(tot) : "-"}}
                    </div>
                  </div>
                </div>
              </td>
            `;
          }}
        }});

        tr.innerHTML = html;
        tbody.appendChild(tr);
      }});

      calculateTotalsCD();
    }}

    function calculateTotalsCD() {{
      const steelTotals = [0, 0, 0, 0, 0];
      const arameTotals = [0, 0, 0, 0, 0];
      const arameQuoted = [false, false, false, false, false];
      const fullTotals = [0, 0, 0, 0, 0];
      const missingActive = [0, 0, 0, 0, 0];

      let activeTotal = 0;
      let activeVolume = 0;
      let activeItemsCount = 0;
      const supplierUsage = [0, 0, 0, 0, 0];

      let bestMixCD = 0;

      itemsData.forEach(item => {{
        if (item.qtd > 0) {{
          activeVolume += item.qtd;
          activeItemsCount++;

          let minP = Infinity;
          item.pricesCD.forEach(p => {{
            if (p !== null && p < minP) minP = p;
          }});
          if (minP < Infinity) bestMixCD += (minP * item.qtd);

          const selIdx = selectedCD[item.id];
          if (selIdx !== undefined && item.pricesCD[selIdx] !== null) {{
            const pSel = item.pricesCD[selIdx];
            activeTotal += (pSel * item.qtd);
            supplierUsage[selIdx] += (pSel * item.qtd);
          }}
        }}

        // Totais de cada coluna
        suppliersCD.forEach((_, sIdx) => {{
          const p = item.pricesCD[sIdx];
          if (p !== null) {{
            const val = p * item.qtd;
            if (item.cat === "arame") {{
              arameTotals[sIdx] += val;
              arameQuoted[sIdx] = true;
            }} else {{
              steelTotals[sIdx] += val;
            }}
            fullTotals[sIdx] += val;
          }} else if (item.qtd > 0) {{
            missingActive[sIdx]++;
          }}
        }});
      }});

      // Atualiza Rodapé da Tabela CD
      suppliersCD.forEach((_, sIdx) => {{
        const subEl = document.getElementById(`subtot-cd-${{sIdx}}`);
        if (subEl) subEl.innerText = formatBRL(steelTotals[sIdx]);

        const arameEl = document.getElementById(`arame-cd-${{sIdx}}`);
        if (arameEl) {{
          arameEl.innerText = arameQuoted[sIdx] ? formatBRL(arameTotals[sIdx]) : "Não cotado";
        }}

        const totEl = document.getElementById(`tot-cd-${{sIdx}}`);
        if (totEl) totEl.innerText = formatBRL(fullTotals[sIdx]);
      }});

      // Atualiza KPIs do Topo se CD estiver ativo
      if (activeTab !== "barra-reta") {{
        document.getElementById("card-custo-total").innerText = formatBRL(activeTotal);
        
        const worstCD = Math.max(...fullTotals); // Kauefer CD (R$ 112.884,65)
        const ecoVal = worstCD - activeTotal;
        const ecoPct = worstCD > 0 ? (ecoVal / worstCD) * 100 : 0;

        document.getElementById("card-economia-potencial").innerText = formatBRL(Math.max(0, ecoVal));
        document.getElementById("card-economia-pct").innerText = `${{formatNum(Math.max(0, ecoPct), 1)}}% de redução vs. maior CD`;

        const diffVsWorst = worstCD > 0 ? ((activeTotal - worstCD) / worstCD) * 100 : 0;
        document.getElementById("card-custo-diff").innerText = `${{formatNum(diffVsWorst, 1)}}% vs. pior proposta CD`;

        document.getElementById("card-volume-total").innerText = `${{formatNum(activeVolume, 2)}} kg`;
        document.getElementById("card-itens-count").innerText = `${{activeItemsCount}} itens com volume ativo`;

        const usedCount = supplierUsage.filter(v => v > 0).length;
        document.getElementById("card-fornecedores-count").innerText = usedCount;
        
        let suppText = "Paranafer + Nortefer";
        if (usedCount > 1) suppText = `Mix (${{usedCount}} Fornecedores CD)`;
        else if (usedCount === 1) {{
          const singleIdx = supplierUsage.findIndex(v => v > 0);
          suppText = `${{suppliersCD[singleIdx].nome}} (100% CD)`;
        }}
        document.getElementById("card-fornecedores-selecionados").innerText = suppText;

        document.getElementById("card-cobertura-pct").innerText = "100% Equalizado";
        document.getElementById("card-status-mapa").innerText = "5 cotações CD + 4 Barra";
      }}

      // Atualiza Resumo do Cenário CD (Card 3 inferior)
      const worstGlobalCD = Math.max(...fullTotals);
      const ecoCD = worstGlobalCD - activeTotal;
      const ecoCDPct = worstGlobalCD > 0 ? (ecoCD / worstGlobalCD) * 100 : 0;

      document.getElementById("summary-custo-total-cd").innerText = formatBRL(activeTotal);
      document.getElementById("summary-custo-menor-cd").innerText = formatBRL(bestMixCD);
      document.getElementById("summary-economia-cd").innerText = `${{formatBRL(ecoCD)}} (-${{formatNum(ecoCDPct, 1)}}%)`;

      updateDonutChartCD(activeTotal, supplierUsage);
      updateCompositionBarCD();
      updateDetailCardCD();
      renderExecutiveRanking();
      renderSiengeTable();
      renderBarCharts();
    }}

    // ==================== LÓGICA DE BARRA RETA ====================
    function handleSearchBR(val) {{
      searchBR = val.toLowerCase().trim();
      renderTableBR();
    }}
    function setFilterBR(cat) {{
      filterCatBR = cat;
      renderTableBR();
    }}
    function toggleOnlyActiveBR(checked) {{
      onlyActiveBR = checked;
      renderTableBR();
    }}
    function selectSingleSupplierBR(sIdx) {{
      itemsData.forEach(item => {{
        if (item.pricesBR[sIdx] !== null) {{
          selectedBR[item.id] = sIdx;
        }}
      }});
      renderTableBR();
    }}
    function selectBestPricesAllBR() {{
      itemsData.forEach(item => {{
        let bestIdx = -1;
        let minP = Infinity;
        item.pricesBR.forEach((p, idx) => {{
          if (p !== null && p < minP) {{
            minP = p;
            bestIdx = idx;
          }}
        }});
        if (bestIdx !== -1) selectedBR[item.id] = bestIdx;
      }});
      renderTableBR();
    }}
    function handleRadioSelectBR(itemId, sIdx) {{
      selectedBR[itemId] = sIdx;
      inspectedItemId = itemId;
      renderTableBR();
    }}

    function renderTableBR() {{
      const tbody = document.getElementById("table-body-br");
      if (!tbody) return;
      tbody.innerHTML = "";

      const filtered = itemsData.filter(item => {{
        if (onlyActiveBR && item.qtd <= 0) return false;
        if (filterCatBR !== "all" && item.cat !== filterCatBR) return false;
        if (searchBR) {{
          const term = searchBR;
          const matchCod = item.cod.toLowerCase().includes(term);
          const matchName = item.nome.toLowerCase().includes(term);
          const matchDesc = item.desc.toLowerCase().includes(term);
          if (!matchCod && !matchName && !matchDesc) return false;
        }}
        return true;
      }});

      filtered.forEach(item => {{
        const isInspected = (item.id === inspectedItemId);
        let minP = Infinity;
        item.pricesBR.forEach(p => {{
          if (p !== null && p < minP) minP = p;
        }});

        const tr = document.createElement("tr");
        tr.className = `hover:bg-slate-50 transition cursor-pointer ${{isInspected ? "bg-blue-50/20" : ""}}`;
        tr.onclick = () => {{
          inspectedItemId = item.id;
          updateDetailCardCD();
          updateDetailCardBR();
        }};

        let html = `
          <td class="py-2.5 px-2.5 font-semibold text-slate-500 font-mono text-[11px] sticky-col-1 bg-white">${{item.cod}}</td>
          <td class="py-2.5 px-2.5 sticky-col-2 bg-white">
            <div class="font-bold text-slate-900">${{item.nome}}</div>
            <div class="text-[10px] text-slate-400 truncate max-w-[205px]">${{item.desc}}</div>
          </td>
          <td class="py-2.5 px-1.5 text-center font-bold text-slate-500">${{item.unid}}</td>
          <td class="py-2.5 px-2.5 text-right">
            <input type="number" step="0.01" value="${{item.qtd.toFixed(2)}}" onchange="updateItemQty(${{item.id}}, this.value)" class="w-20 py-1 px-1.5 text-right font-bold text-slate-800 bg-slate-50 hover:bg-slate-100 focus:bg-white border border-slate-200 rounded-lg text-xs num-tabular focus:outline-none focus:ring-1 focus:ring-[#0a2540]">
          </td>
        `;

        // 4 Colunas Fornecedores Barra Reta
        suppliersBR.forEach((_, sIdx) => {{
          const p = item.pricesBR[sIdx];
          const isSelected = (selectedBR[item.id] === sIdx);
          const isBest = (p !== null && p === minP);
          const tot = (p !== null && item.qtd > 0) ? (p * item.qtd) : 0;

          let diffBadge = "";
          if (p !== null && minP < Infinity && p > minP) {{
            const dVal = p - minP;
            const dPct = (dVal / minP) * 100;
            diffBadge = `<span class="diff-pill diff-pill-higher">+${{dPct.toFixed(1)}}%</span>`;
          }} else if (isBest) {{
            diffBadge = `<span class="diff-pill diff-pill-best">Menor</span>`;
          }}

          if (p === null) {{
            html += `
              <td class="py-2.5 px-2.5 text-center border-l border-slate-100 bg-slate-50/40 text-slate-400 italic text-[11px]">
                Não cotado
              </td>
            `;
          }} else {{
            html += `
              <td class="py-2.5 px-2.5 border-l border-slate-100 ${{isSelected ? "cell-selected" : ""}} ${{isBest ? "cell-best" : ""}}" onclick="handleRadioSelectBR(${{item.id}}, ${{sIdx}})">
                <div class="flex items-center justify-between">
                  <div class="saas-radio ${{isSelected ? "checked" : ""}}"></div>
                  <div class="text-right flex-1 ml-1.5">
                    <div class="flex items-center justify-end space-x-1">
                      <span class="font-extrabold text-slate-900 num-tabular text-xs">${{formatBRL(p)}}</span>
                      ${{diffBadge}}
                    </div>
                    <div class="text-[10px] text-slate-500 num-tabular mt-0.5">
                      ${{item.qtd > 0 ? formatBRL(tot) : "-"}}
                    </div>
                  </div>
                </div>
              </td>
            `;
          }}
        }});

        tr.innerHTML = html;
        tbody.appendChild(tr);
      }});

      calculateTotalsBR();
    }}

    function calculateTotalsBR() {{
      const steelTotals = [0, 0, 0, 0];
      const arameTotals = [0, 0, 0, 0];
      const fullTotals = [0, 0, 0, 0];

      let activeTotal = 0;
      let activeVolume = 0;
      let activeItemsCount = 0;
      const supplierUsage = [0, 0, 0, 0];

      let bestMixBR = 0;

      itemsData.forEach(item => {{
        if (item.qtd > 0) {{
          activeVolume += item.qtd;
          activeItemsCount++;

          let minP = Infinity;
          item.pricesBR.forEach(p => {{
            if (p !== null && p < minP) minP = p;
          }});
          if (minP < Infinity) bestMixBR += (minP * item.qtd);

          const selIdx = selectedBR[item.id];
          if (selIdx !== undefined && item.pricesBR[selIdx] !== null) {{
            const pSel = item.pricesBR[selIdx];
            activeTotal += (pSel * item.qtd);
            supplierUsage[selIdx] += (pSel * item.qtd);
          }}
        }}

        // Totais de cada coluna
        suppliersBR.forEach((_, sIdx) => {{
          const p = item.pricesBR[sIdx];
          if (p !== null) {{
            const val = p * item.qtd;
            if (item.cat === "arame") {{
              arameTotals[sIdx] += val;
            }} else {{
              steelTotals[sIdx] += val;
            }}
            fullTotals[sIdx] += val;
          }}
        }});
      }});

      // Atualiza Rodapé da Tabela BR
      suppliersBR.forEach((_, sIdx) => {{
        const subEl = document.getElementById(`subtot-br-${{sIdx}}`);
        if (subEl) subEl.innerText = formatBRL(steelTotals[sIdx]);

        const arameEl = document.getElementById(`arame-br-${{sIdx}}`);
        if (arameEl) arameEl.innerText = formatBRL(arameTotals[sIdx]);

        const totEl = document.getElementById(`tot-br-${{sIdx}}`);
        if (totEl) totEl.innerText = formatBRL(fullTotals[sIdx]);
      }});

      // Atualiza KPIs do Topo se Barra Reta estiver ativa
      if (activeTab === "barra-reta") {{
        document.getElementById("card-custo-total").innerText = formatBRL(activeTotal);
        
        const worstBR = Math.max(...fullTotals); // NORTEFER Barra (R$ 103.588,67)
        const ecoVal = worstBR - activeTotal;
        const ecoPct = worstBR > 0 ? (ecoVal / worstBR) * 100 : 0;

        document.getElementById("card-economia-potencial").innerText = formatBRL(Math.max(0, ecoVal));
        document.getElementById("card-economia-pct").innerText = `${{formatNum(Math.max(0, ecoPct), 1)}}% de redução vs. maior Barra`;

        const diffVsWorst = worstBR > 0 ? ((activeTotal - worstBR) / worstBR) * 100 : 0;
        document.getElementById("card-custo-diff").innerText = `${{formatNum(diffVsWorst, 1)}}% vs. pior proposta Barra`;

        document.getElementById("card-volume-total").innerText = `${{formatNum(activeVolume, 2)}} kg`;
        document.getElementById("card-itens-count").innerText = `${{activeItemsCount}} itens com volume ativo`;

        const usedCount = supplierUsage.filter(v => v > 0).length;
        document.getElementById("card-fornecedores-count").innerText = usedCount;

        let suppText = "ArcelorMittal (100% Barra)";
        if (usedCount > 1) suppText = `Mix (${{usedCount}} Fornecedores Barra)`;
        else if (usedCount === 1) {{
          const singleIdx = supplierUsage.findIndex(v => v > 0);
          suppText = `${{suppliersBR[singleIdx].nome}} (100% Barra)`;
        }}
        document.getElementById("card-fornecedores-selecionados").innerText = suppText;

        document.getElementById("card-cobertura-pct").innerText = "100% Equalizado";
        document.getElementById("card-status-mapa").innerText = "4 usinas/distribuidoras Barra";
      }}

      // Atualiza Resumo do Cenário BR (Card 3 inferior)
      const worstGlobalBR = Math.max(...fullTotals);
      const ecoBR = worstGlobalBR - activeTotal;
      const ecoBRPct = worstGlobalBR > 0 ? (ecoBR / worstGlobalBR) * 100 : 0;

      document.getElementById("summary-custo-total-br").innerText = formatBRL(activeTotal);
      document.getElementById("summary-custo-menor-br").innerText = formatBRL(bestMixBR);
      document.getElementById("summary-economia-br").innerText = `${{formatBRL(ecoBR)}} (-${{formatNum(ecoBRPct, 1)}}%)`;

      updateDonutChartBR(activeTotal, supplierUsage);
      updateCompositionBarBR();
      updateDetailCardBR();
    }}

    // ATUALIZA QUANTIDADE DE QUALQUER ITEM
    function updateItemQty(id, val) {{
      const num = parseFloat(val) || 0;
      const item = itemsData.find(i => i.id === id);
      if (item) {{
        item.qtd = num;
        renderTableCD();
        renderTableBR();
        runSim();
      }}
    }}

    // RESTAURAR VOLUME OFICIAL ZOCCO R01 (18.999,00 KG)
    function resetToOfficialVolume() {{
      const official = {{ 1: 1687.0, 2: 0.0, 3: 0.0, 4: 518.0, 5: 12672.0, 6: 3722.0, 7: 0.0, 8: 0.0, 9: 400.0 }};
      itemsData.forEach(item => {{
        item.qtd = official[item.id] !== undefined ? official[item.id] : 0.0;
      }});
      renderTableCD();
      renderTableBR();
      runSim();
      alert("Volumes oficiais da etapa (18.999,00 kg / 19,00 t) restaurados com sucesso!");
    }}

    // ==================== DONUT CHARTS ====================
    function updateDonutChartCD(totalVal, usageArr) {{
      const svg = document.getElementById("donut-svg-cd");
      const legend = document.getElementById("donut-legend-cd");
      const centerEl = document.getElementById("donut-center-val-cd");
      if (!svg || !legend || !centerEl) return;

      centerEl.innerText = formatBRL(totalVal).replace(",00", "");

      const radius = 38;
      const circumference = 2 * Math.PI * radius;
      let offset = 0;

      svg.querySelectorAll(".donut-segment").forEach(e => e.remove());
      legend.innerHTML = "";

      suppliersCD.forEach((s, idx) => {{
        const val = usageArr[idx] || 0;
        const pct = totalVal > 0 ? (val / totalVal) : 0;
        const strokeDasharray = `${{pct * circumference}} ${{circumference}}`;

        if (val > 0) {{
          const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
          circle.setAttribute("class", "donut-segment");
          circle.setAttribute("cx", "50");
          circle.setAttribute("cy", "50");
          circle.setAttribute("r", radius);
          circle.setAttribute("stroke", s.color);
          circle.setAttribute("stroke-width", "14");
          circle.setAttribute("fill", "transparent");
          circle.setAttribute("stroke-dasharray", strokeDasharray);
          circle.setAttribute("stroke-dashoffset", -offset);
          svg.appendChild(circle);

          offset += pct * circumference;
        }}

        // Legenda
        const row = document.createElement("div");
        row.className = "flex items-center justify-between text-[11px]";
        row.innerHTML = `
          <div class="flex items-center space-x-1.5">
            <span class="w-2.5 h-2.5 rounded-full" style="background-color: ${{s.color}}"></span>
            <span class="font-semibold text-slate-700">${{s.nome}}</span>
          </div>
          <div class="num-tabular font-bold text-slate-800">
            ${{formatBRL(val)}} <span class="text-slate-400 font-normal">(${{(pct * 100).toFixed(1)}}%)</span>
          </div>
        `;
        legend.appendChild(row);
      }});
    }}

    function updateDonutChartBR(totalVal, usageArr) {{
      const svg = document.getElementById("donut-svg-br");
      const legend = document.getElementById("donut-legend-br");
      const centerEl = document.getElementById("donut-center-val-br");
      if (!svg || !legend || !centerEl) return;

      centerEl.innerText = formatBRL(totalVal).replace(",00", "");

      const radius = 38;
      const circumference = 2 * Math.PI * radius;
      let offset = 0;

      svg.querySelectorAll(".donut-segment").forEach(e => e.remove());
      legend.innerHTML = "";

      suppliersBR.forEach((s, idx) => {{
        const val = usageArr[idx] || 0;
        const pct = totalVal > 0 ? (val / totalVal) : 0;
        const strokeDasharray = `${{pct * circumference}} ${{circumference}}`;

        if (val > 0) {{
          const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
          circle.setAttribute("class", "donut-segment");
          circle.setAttribute("cx", "50");
          circle.setAttribute("cy", "50");
          circle.setAttribute("r", radius);
          circle.setAttribute("stroke", s.color);
          circle.setAttribute("stroke-width", "14");
          circle.setAttribute("fill", "transparent");
          circle.setAttribute("stroke-dasharray", strokeDasharray);
          circle.setAttribute("stroke-dashoffset", -offset);
          svg.appendChild(circle);

          offset += pct * circumference;
        }}

        // Legenda
        const row = document.createElement("div");
        row.className = "flex items-center justify-between text-[11px]";
        row.innerHTML = `
          <div class="flex items-center space-x-1.5">
            <span class="w-2.5 h-2.5 rounded-full" style="background-color: ${{s.color}}"></span>
            <span class="font-semibold text-slate-700">${{s.nome}}</span>
          </div>
          <div class="num-tabular font-bold text-slate-800">
            ${{formatBRL(val)}} <span class="text-slate-400 font-normal">(${{(pct * 100).toFixed(1)}}%)</span>
          </div>
        `;
        legend.appendChild(row);
      }});
    }}

    // ==================== BARRAS DE COMPOSIÇÃO DE VOLUME ====================
    function updateCompositionBarCD() {{
      const bar = document.getElementById("composition-bar-cd");
      const leg = document.getElementById("composition-legend-cd");
      if (!bar || !leg) return;

      const totalVol = itemsData.reduce((acc, i) => acc + i.qtd, 0);
      bar.innerHTML = "";
      leg.innerHTML = "";

      const colors = ["#0891b2", "#0284c7", "#3b82f6", "#1d4ed8", "#475569"];
      const activeItems = itemsData.filter(i => i.qtd > 0);

      activeItems.forEach((item, idx) => {{
        const pct = totalVol > 0 ? (item.qtd / totalVol) * 100 : 0;
        const c = colors[idx % colors.length];

        const seg = document.createElement("div");
        seg.style.width = `${{pct}}%`;
        seg.style.backgroundColor = c;
        seg.className = "h-full transition-all";
        seg.title = `${{item.nome}}: ${{pct.toFixed(1)}}%`;
        bar.appendChild(seg);
      }});

      leg.innerHTML = `
        <span>5.0mm (8,9%)</span>
        <span>10.0mm (2,7%)</span>
        <span class="font-bold text-cyan-900">12.5mm (66,7%)</span>
        <span>16.0mm (19,6%)</span>
        <span>Arame (2,1%)</span>
      `;
    }}

    function updateCompositionBarBR() {{
      const bar = document.getElementById("composition-bar-br");
      const leg = document.getElementById("composition-legend-br");
      if (!bar || !leg) return;

      const totalVol = itemsData.reduce((acc, i) => acc + i.qtd, 0);
      bar.innerHTML = "";
      leg.innerHTML = "";

      const colors = ["#1d4ed8", "#2563eb", "#3b82f6", "#60a5fa", "#64748b"];
      const activeItems = itemsData.filter(i => i.qtd > 0);

      activeItems.forEach((item, idx) => {{
        const pct = totalVol > 0 ? (item.qtd / totalVol) * 100 : 0;
        const c = colors[idx % colors.length];

        const seg = document.createElement("div");
        seg.style.width = `${{pct}}%`;
        seg.style.backgroundColor = c;
        seg.className = "h-full transition-all";
        seg.title = `${{item.nome}}: ${{pct.toFixed(1)}}%`;
        bar.appendChild(seg);
      }});

      leg.innerHTML = `
        <span>5.0mm (8,9%)</span>
        <span>10.0mm (2,7%)</span>
        <span class="font-bold text-blue-900">12.5mm (66,7%)</span>
        <span>16.0mm (19,6%)</span>
        <span>Arame (2,1%)</span>
      `;
    }}

    // ==================== DETALHE DO ITEM INSPECIONADO ====================
    function updateDetailCardCD() {{
      const item = itemsData.find(i => i.id === inspectedItemId) || itemsData[4];
      if (!item) return;

      const codEl = document.getElementById("detail-item-cod-cd");
      const catEl = document.getElementById("detail-item-cat-cd");
      const nameEl = document.getElementById("detail-item-name-cd");
      const descEl = document.getElementById("detail-item-desc-cd");
      const volEl = document.getElementById("detail-item-vol-cd");
      const priceEl = document.getElementById("detail-item-price-cd");
      const totEl = document.getElementById("detail-item-total-cd");

      if (codEl) codEl.innerText = item.cod;
      if (catEl) catEl.innerText = item.cat.toUpperCase();
      if (nameEl) nameEl.innerText = item.nome;
      if (descEl) descEl.innerText = item.desc;
      if (volEl) volEl.innerText = `${{formatNum(item.qtd, 2)}} kg`;

      const selIdx = selectedCD[item.id] !== undefined ? selectedCD[item.id] : 0;
      const p = item.pricesCD[selIdx] || 0;
      if (priceEl) priceEl.innerText = `${{formatBRL(p)}}/kg (${{suppliersCD[selIdx].nome}})`;
      if (totEl) totEl.innerText = formatBRL(p * item.qtd);
    }}

    function updateDetailCardBR() {{
      const item = itemsData.find(i => i.id === inspectedItemId) || itemsData[4];
      if (!item) return;

      const codEl = document.getElementById("detail-item-cod-br");
      const catEl = document.getElementById("detail-item-cat-br");
      const nameEl = document.getElementById("detail-item-name-br");
      const descEl = document.getElementById("detail-item-desc-br");
      const volEl = document.getElementById("detail-item-vol-br");
      const priceEl = document.getElementById("detail-item-price-br");
      const totEl = document.getElementById("detail-item-total-br");

      if (codEl) codEl.innerText = item.cod;
      if (catEl) catEl.innerText = item.cat.toUpperCase();
      if (nameEl) nameEl.innerText = item.nome;
      if (descEl) descEl.innerText = item.desc;
      if (volEl) volEl.innerText = `${{formatNum(item.qtd, 2)}} kg`;

      const selIdx = selectedBR[item.id] !== undefined ? selectedBR[item.id] : 0;
      const p = item.pricesBR[selIdx] || 0;
      if (priceEl) priceEl.innerText = `${{formatBRL(p)}}/kg (${{suppliersBR[selIdx].nome}})`;
      if (totEl) totEl.innerText = formatBRL(p * item.qtd);
    }}

    // ==================== RANKING EXECUTIVO GERAL (9 PROPOSTAS) ====================
    function renderExecutiveRanking() {{
      const tbody = document.getElementById("executive-ranking-body");
      if (!tbody) return;
      tbody.innerHTML = "";

      const calcTotal = (isCD, sIdx) => {{
        let tot = 0;
        let missing = 0;
        itemsData.forEach(i => {{
          if (i.qtd > 0) {{
            const p = isCD ? i.pricesCD[sIdx] : i.pricesBR[sIdx];
            if (p !== null) tot += p * i.qtd;
            else missing++;
          }}
        }});
        return {{ tot, missing }};
      }};

      const cd0 = calcTotal(true, 0); // Paranafer CD
      const cd1 = calcTotal(true, 1); // Arcelor CD
      const cd2 = calcTotal(true, 2); // Gerdau CD
      const cd3 = calcTotal(true, 3); // Kauefer CD
      const cd4 = calcTotal(true, 4); // NORTEFER CD

      const br0 = calcTotal(false, 0); // Arcelor Barra
      const br1 = calcTotal(false, 1); // Gerdau Barra
      const br2 = calcTotal(false, 2); // Kauefer Barra
      const br3 = calcTotal(false, 3); // NORTEFER Barra

      const canteiroAddon = 17047.00; // R$ 5.888 sucata 6,5% + R$ 11.159 M.O. bancada

      const list = [
        {{
          "nome": "ArcelorMittal",
          "modalidade": "Barra Reta (12m)",
          "badgeMod": "bg-slate-100 text-slate-700",
          "totalNominal": br0.tot,
          "custoReal": br0.tot + canteiroAddon,
          "prazo": "Faturado Usina • Completo",
          "obs": "1º Líder em Barra Reta (menor NF nominal: R$ 93.787,27). Porém exige bancada de corte no canteiro e gera 6,5% de perda em sucata."
        }},
        {{
          "nome": "Gerdau",
          "modalidade": "Barra Reta (12m)",
          "badgeMod": "bg-slate-100 text-slate-700",
          "totalNominal": br1.tot,
          "custoReal": br1.tot + canteiroAddon,
          "prazo": "Antecipado B014 • Completo",
          "obs": "2º Lugar em Barra Reta (100% completo após W10234125: 6,3/8mm a R$ 5,23/kg, 12,5/16/20/25mm a R$ 4,84/kg e Arame a R$ 10,63/kg)."
        }},
        {{
          "nome": "Paranafer",
          "modalidade": "Corte e Dobra",
          "badgeMod": "bg-cyan-100 text-cyan-900 font-bold",
          "totalNominal": cd0.tot,
          "custoReal": cd0.tot + 3240.00,
          "prazo": "30D Boleto • S/ Arame",
          "isRecommended": true,
          "obs": "1º Líder no Aço em Corte e Dobra (R$ 101.447,64 a R$ 5,42/kg no CA-50 e R$ 5,80/kg no CA-60). Comprando o arame à parte (R$ 3.240), fecha o menor pacote real da obra em R$ 104.687,64."
        }},
        {{
          "nome": "Kauefer",
          "modalidade": "Barra Reta (12m)",
          "badgeMod": "bg-slate-100 text-slate-700",
          "totalNominal": br2.tot,
          "custoReal": br2.tot + canteiroAddon,
          "prazo": "Faturado • Completo",
          "obs": "3º Lugar em Barra Reta conforme planilha ODS (R$ 5,15/kg no CA-50 10 a 25mm e R$ 11,05/kg no arame)."
        }},
        {{
          "nome": "NORTEFER",
          "modalidade": "Barra Reta (12m)",
          "badgeMod": "bg-slate-100 text-slate-700",
          "totalNominal": br3.tot,
          "custoReal": br3.tot + canteiroAddon,
          "prazo": "Faturado • Completo",
          "obs": "4º Lugar em Barra Reta (Barra Corte Sob Medida: CA-60 5mm a R$ 6,05/kg, CA-50 10/12,5/16mm a R$ 5,33/kg e 20/25mm a R$ 5,23/kg)."
        }},
        {{
          "nome": "ArcelorMittal (CD)",
          "modalidade": "Corte e Dobra",
          "badgeMod": "bg-cyan-100 text-cyan-900 font-bold",
          "totalNominal": cd1.tot,
          "custoReal": cd1.tot,
          "prazo": "15 Dias Corridos • Completo",
          "obs": "1º Lugar entre as propostas 100% completas de Corte e Dobra (inclui aço + arame a R$ 8,11/kg). Serviço Belgo Pronto com entrega em 15 dias."
        }},
        {{
          "nome": "Gerdau (CD)",
          "modalidade": "Corte e Dobra",
          "badgeMod": "bg-cyan-100 text-cyan-900 font-bold",
          "totalNominal": cd2.tot,
          "custoReal": cd2.tot,
          "prazo": "Antecipado B014 • Completo",
          "obs": "2º Lugar entre as propostas completas de Corte e Dobra (R$ 106.505,14 no aço + R$ 4.251,30 no arame RL50kg)."
        }},
        {{
          "nome": "NORTEFER (CD)",
          "modalidade": "Corte e Dobra",
          "badgeMod": "bg-cyan-100 text-cyan-900 font-bold",
          "totalNominal": cd4.tot,
          "custoReal": cd4.tot,
          "prazo": "Faturado Programado • Completo",
          "obs": "3º Lugar entre as propostas completas de Corte e Dobra após atualização formal de 22/10 (CA-60 5mm e CA-50 6,3/8mm a R$ 6,05/kg; CA-50 10 a 25mm a R$ 5,85/kg)."
        }},
        {{
          "nome": "Kauefer (CD)",
          "modalidade": "Corte e Dobra",
          "badgeMod": "bg-cyan-100 text-cyan-900 font-bold",
          "totalNominal": cd3.tot,
          "custoReal": cd3.tot,
          "prazo": "Faturado • Completo",
          "obs": "4º Lugar entre as propostas completas de Corte e Dobra (CA-50 10 a 25mm a R$ 5,80/kg e Arame a R$ 11,05/kg)."
        }}
      ];

      list.sort((a, b) => a.totalNominal - b.totalNominal);
      const minGlobal = list[0].totalNominal;

      list.forEach((item, idx) => {{
        const diff = item.totalNominal - minGlobal;
        const diffPct = minGlobal > 0 ? (diff / minGlobal) * 100 : 0;
        const diffText = (idx === 0)
          ? '<span class="text-emerald-700 font-bold">Menor Valor Nominal</span>'
          : `<span class="text-slate-600 font-semibold">+${{formatBRL(diff)}} (+${{formatNum(diffPct, 1)}}%)</span>`;

        const tr = document.createElement("tr");
        if (item.isRecommended) tr.className = "bg-cyan-50/60 font-semibold";
        else if (idx === 0) tr.className = "bg-slate-50 font-semibold";

        tr.innerHTML = `
          <td class="py-3 px-3 text-center font-bold text-slate-700">${{idx + 1}}º</td>
          <td class="py-3 px-3 font-bold text-slate-900">
            ${{item.nome}}
            ${{item.isRecommended ? '<span class="ml-1.5 px-2 py-0.5 rounded text-[9px] bg-emerald-100 text-emerald-800 font-black">Melhor Custo Real CD</span>' : ''}}
            ${{item.nome === "ArcelorMittal (CD)" ? '<span class="ml-1.5 px-2 py-0.5 rounded text-[9px] bg-blue-100 text-blue-800 font-black">1º CD Completo</span>' : ''}}
          </td>
          <td class="py-3 px-3 text-center"><span class="px-2.5 py-0.5 rounded text-[10px] ${{item.badgeMod}}">${{item.modalidade}}</span></td>
          <td class="py-3 px-3 text-right font-black text-slate-900 num-tabular">${{formatBRL(item.totalNominal)}}</td>
          <td class="py-3 px-3 text-right num-tabular">${{diffText}}</td>
          <td class="py-3 px-3 text-right font-bold ${{item.isRecommended ? 'text-emerald-700 font-black' : 'text-slate-700'}} num-tabular">${{formatBRL(item.custoReal)}}</td>
          <td class="py-3 px-3 text-center text-[11px] font-semibold">${{item.prazo}}</td>
          <td class="py-3 px-3 text-slate-600">${{item.obs}}</td>
        `;
        tbody.appendChild(tr);
      }});
    }}

    // ==================== SIMULADOR DE CONTRAPROPOSTA ====================
    function setFastSim(pct) {{
      document.getElementById("sim-disc-input").value = pct.toFixed(1);
      document.getElementById("sim-disc-slider").value = pct.toFixed(1);
      runSim();
    }}
    function runSimFromSlider(val) {{
      document.getElementById("sim-disc-input").value = parseFloat(val).toFixed(1);
      runSim();
    }}
    function runSimFromInput(val) {{
      document.getElementById("sim-disc-slider").value = parseFloat(val) || 0;
      runSim();
    }}

    function runSim() {{
      const suppKey = document.getElementById("sim-supp-select").value;
      const disc = parseFloat(document.getElementById("sim-disc-input").value) || 0;
      const mult = Math.max(0, 1 - (disc / 100));

      const isCD = suppKey.startsWith("cd-");
      const sIdx = parseInt(suppKey.split("-")[1]);

      let totalOrig = 0;
      let totalSim = 0;
      const simRows = [];

      itemsData.forEach(item => {{
        if (item.qtd > 0) {{
          const pOrig = isCD ? item.pricesCD[sIdx] : item.pricesBR[sIdx];
          if (pOrig !== null) {{
            const pSim = pOrig * mult;
            const totOrig = pOrig * item.qtd;
            const totSim = pSim * item.qtd;
            const ecoItem = totOrig - totSim;

            totalOrig += totOrig;
            totalSim += totSim;

            simRows.push({{
              "cod": item.cod,
              "nome": item.nome,
              "unid": item.unid,
              "qtd": item.qtd,
              "pOrig": pOrig,
              "pSim": pSim,
              "totOrig": totOrig,
              "totSim": totSim,
              "eco": ecoItem
            }});
          }}
        }}
      }});

      const totalEco = totalOrig - totalSim;

      document.getElementById("sim-val-original").innerText = formatBRL(totalOrig);
      document.getElementById("sim-val-desconto").innerText = `-${{formatBRL(totalEco)}} (-${{disc.toFixed(1)}}%)`;
      document.getElementById("sim-val-final").innerText = formatBRL(totalSim);

      const tbody = document.getElementById("sim-table-body");
      if (tbody) {{
        tbody.innerHTML = "";
        simRows.forEach(r => {{
          const tr = document.createElement("tr");
          tr.className = "hover:bg-slate-50 transition";
          tr.innerHTML = `
            <td class="py-2.5 px-3 font-semibold text-slate-500 font-mono text-[11px]">${{r.cod}}</td>
            <td class="py-2.5 px-3 font-bold text-slate-900">${{r.nome}}</td>
            <td class="py-2.5 px-2 text-center text-slate-500">${{r.unid}}</td>
            <td class="py-2.5 px-3 text-right font-bold text-slate-700 num-tabular">${{formatNum(r.qtd, 2)}}</td>
            <td class="py-2.5 px-3 text-right text-slate-500 num-tabular">${{formatBRL(r.pOrig)}}</td>
            <td class="py-2.5 px-3 text-right bg-emerald-50 text-emerald-950 font-black num-tabular">${{formatBRL(r.pSim)}}</td>
            <td class="py-2.5 px-3 text-right text-slate-500 num-tabular">${{formatBRL(r.totOrig)}}</td>
            <td class="py-2.5 px-3 text-right bg-emerald-50 text-emerald-950 font-black num-tabular">${{formatBRL(r.totSim)}}</td>
            <td class="py-2.5 px-3 text-right text-emerald-700 font-bold num-tabular">-${{formatBRL(r.eco)}}</td>
          `;
          tbody.appendChild(tr);
        }});
      }}

      const footOrig = document.getElementById("sim-foot-tot-orig");
      const footNeg = document.getElementById("sim-foot-tot-neg");
      const footEco = document.getElementById("sim-foot-tot-eco");
      if (footOrig) footOrig.innerText = formatBRL(totalOrig);
      if (footNeg) footNeg.innerText = formatBRL(totalSim);
      if (footEco) footEco.innerText = `-${{formatBRL(totalEco)}}`;
    }}

    function copyNegotiatedProposal() {{
      const suppKey = document.getElementById("sim-supp-select").value;
      const disc = parseFloat(document.getElementById("sim-disc-input").value) || 0;
      const mult = Math.max(0, 1 - (disc / 100));

      const isCD = suppKey.startsWith("cd-");
      const sIdx = parseInt(suppKey.split("-")[1]);
      const suppObj = isCD ? suppliersCD[sIdx] : suppliersBR[sIdx];

      let msg = `*CONTRAPROPOSTA COMERCIAL - OBRA 2002 GIARDINI (EVENSE)*\\n`;
      msg += `Fornecedor: ${{suppObj.nome}} (${{suppObj.modalidade}})\\n`;
      msg += `Volume Total da Etapa: 18.999,00 kg (19,00 t) | Desconto Solicitado: ${{disc.toFixed(1)}}%\\n\\n`;
      msg += `*Bitolas e Preços Unitários Negociados:*\\n`;

      let totalSim = 0;
      itemsData.forEach(item => {{
        if (item.qtd > 0) {{
          const pOrig = isCD ? item.pricesCD[sIdx] : item.pricesBR[sIdx];
          if (pOrig !== null) {{
            const pSim = pOrig * mult;
            totalSim += (pSim * item.qtd);
            msg += `- ${{item.nome}}: de ${{formatBRL(pOrig)}} por *${{formatBRL(pSim)}}/kg* (Vol: ${{formatNum(item.qtd, 2)}} kg)\\n`;
          }}
        }}
      }});

      msg += `\\n*Valor Total da Contraproposta: ${{formatBRL(totalSim)}}*\\n`;
      msg += `Condição de Pagamento: ${{suppObj.prazo}}\\n`;
      msg += `Frete: Incluso CIF na obra (Av. Londrina, 494 - Zona 08, Maringá - PR)\\n`;
      msg += `Contato da Engenharia: Eng. Marcio Gracional (Eng. de Planejamento / Suprimentos)`;

      navigator.clipboard.writeText(msg).then(() => {{
        const btn = document.getElementById("btn-copy-sim");
        if (btn) {{
          btn.innerHTML = `<svg class="w-4 h-4 text-emerald-400" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12.75l6 6 9-13.5"/></svg><span>Contraproposta Copiada!</span>`;
          setTimeout(() => {{
            btn.innerHTML = `<svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M15.75 17.25v3.375c0 .621-.504 1.125-1.125 1.125h-9.75a1.125 1.125 0 01-1.125-1.125V7.875c0-.621.504-1.125 1.125-1.125H6.75a9.06 9.06 0 011.5.124m7.5 10.376h3.375c.621 0 1.125-.504 1.125-1.125V11.25c0-4.46-3.243-8.161-7.5-8.876a9.06 9.06 0 00-1.5-.124H9.375c-.621 0-1.125.504-1.125 1.125v3.5m7.5 10.375H9.375a1.125 1.125 0 01-1.125-1.125v-9.25m12 6.625v-1.875a3.375 3.375 0 00-3.375-3.375h-1.5a1.125 1.125 0 01-1.125-1.125v-1.5a3.375 3.375 0 00-3.375-3.375H9.75"/></svg><span>Copiar Contraproposta</span>`;
          }}, 2500);
        }}
      }}).catch(() => {{
        alert("Pressione Ctrl+C para copiar.");
      }});
    }}

    // ==================== ANÁLISE TÉCNICA E AUDITORIA ====================
    function renderAuditTable() {{
      const tbody = document.getElementById("audit-table-body");
      if (!tbody) return;
      tbody.innerHTML = "";

      const auditData = [
        {{ "elem": "Estribos de Vigas e Pilares", "bitola": "CA-60 5,0mm", "req": "NBR 7480 / NBR 6118 (fyk &ge; 600 MPa)", "cot": "5,0mm homologado (Menor CD: R$ 5,80 Paranafer | Barra: R$ 5,48 Arcelor)", "status": "Conforme", "parecer": "Bitola padrão para estribamento contínuo. Raio de dobramento conforme norma." }},
        {{ "elem": "Armadura Longitudinal Secundária", "bitola": "CA-50 10,0mm", "req": "NBR 7480 (fyk &ge; 500 MPa, Along. &ge; 8%)", "cot": "10,0mm nervurado (Menor CD: R$ 5,42 Paranafer | Barra: R$ 4,90 Arcelor)", "status": "Conforme", "parecer": "Ensaios mecânicos de tração e escoamento exigidos no certificado de corrida." }},
        {{ "elem": "Armadura Mestre Estrutural (66,7% do peso)", "bitola": "CA-50 12,5mm", "req": "NBR 7480 (Tolerância linear &plusmn; 4,0%)", "cot": "12,5mm nervurado (Menor CD: R$ 5,42 Paranafer | Barra: R$ 4,80 Arcelor)", "status": "Conforme", "parecer": "Bitola principal de pilares e vigas. Corte e Dobra CNC garante precisão nos nós." }},
        {{ "elem": "Armadura Pesada de Transição", "bitola": "CA-50 16,0mm", "req": "NBR 7480 (Dobra a 180° sem trincas)", "cot": "16,0mm nervurado (Menor CD: R$ 5,42 Paranafer | Barra: R$ 4,80 Arcelor)", "status": "Conforme", "parecer": "Dobra mecânica de fábrica elimina microfissuras observadas em dobra manual." }},
        {{ "elem": "Amarração de Armaduras", "bitola": "Arame BWG 16", "req": "Arame recozido macio e flexível (400,00 kg)", "cot": "BWG 16 NF (R$ 8,10/kg Nortefer | R$ 8,11/kg Arcelor | R$ 10,63/kg Gerdau c/ IPI)", "status": "Conforme", "parecer": "Resistência ideal à tração de amarração sem ruptura prematura." }}
      ];

      auditData.forEach(row => {{
        const tr = document.createElement("tr");
        tr.className = "hover:bg-slate-50 transition";
        tr.innerHTML = `
          <td class="py-2.5 px-3">
            <div class="font-bold text-slate-800">${{row.elem}}</div>
            <div class="text-[10px] text-slate-500 font-semibold">${{row.bitola}}</div>
          </td>
          <td class="py-2.5 px-3 text-slate-600">${{row.req}}</td>
          <td class="py-2.5 px-3 font-semibold text-slate-800">${{row.cot}}</td>
          <td class="py-2.5 px-2 text-center">
            <span class="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800">${{row.status}}</span>
          </td>
          <td class="py-2.5 px-3 text-slate-600 text-[11px]">${{row.parecer}}</td>
        `;
        tbody.appendChild(tr);
      }});
    }}

    // ==================== INTEGRAÇÃO SIENGE ERP ====================
    function renderSiengeTable() {{
      const tbody = document.getElementById("sienge-table-body");
      if (!tbody) return;
      tbody.innerHTML = "";

      const mode = document.getElementById("sienge-modality-select") ? document.getElementById("sienge-modality-select").value : "cd";
      const isCD = (mode === "cd");

      let totalVal = 0;
      itemsData.forEach(item => {{
        if (item.qtd > 0) {{
          const suppIdx = isCD ? selectedCD[item.id] : selectedBR[item.id];
          const suppArr = isCD ? suppliersCD : suppliersBR;
          const prices = isCD ? item.pricesCD : item.pricesBR;

          const p = (suppIdx !== undefined && prices[suppIdx] !== null) ? prices[suppIdx] : 0;
          const suppName = (suppIdx !== undefined) ? suppArr[suppIdx].nome : "-";
          const tot = p * item.qtd;
          totalVal += tot;

          const tr = document.createElement("tr");
          tr.className = "hover:bg-slate-50 transition";
          tr.innerHTML = `
            <td class="py-2.5 px-3 font-mono font-bold text-slate-800 text-[11px]">${{item.cod}}</td>
            <td class="py-2.5 px-3 font-semibold text-slate-900">${{item.nome}}</td>
            <td class="py-2.5 px-2 text-center text-slate-500">${{item.unid}}</td>
            <td class="py-2.5 px-3 text-right font-bold text-slate-800 num-tabular">${{formatNum(item.qtd, 2)}}</td>
            <td class="py-2.5 px-3 text-slate-700 font-semibold">${{suppName}}</td>
            <td class="py-2.5 px-3 text-right font-black text-[#0a2540] num-tabular">${{formatBRL(p)}}</td>
            <td class="py-2.5 px-3 text-right font-black text-slate-900 num-tabular">${{formatBRL(tot)}}</td>
          `;
          tbody.appendChild(tr);
        }}
      }});
    }}

    function copySiengeData() {{
      const mode = document.getElementById("sienge-modality-select") ? document.getElementById("sienge-modality-select").value : "cd";
      const isCD = (mode === "cd");

      let tsv = "Código\\tDescrição\\tUnidade\\tQuantidade\\tFornecedor\\tPreço Unitário\\tValor Total\\n";
      itemsData.forEach(item => {{
        if (item.qtd > 0) {{
          const suppIdx = isCD ? selectedCD[item.id] : selectedBR[item.id];
          const suppArr = isCD ? suppliersCD : suppliersBR;
          const prices = isCD ? item.pricesCD : item.pricesBR;
          const p = (suppIdx !== undefined && prices[suppIdx] !== null) ? prices[suppIdx] : 0;
          const suppName = (suppIdx !== undefined) ? suppArr[suppIdx].nome : "-";
          const tot = p * item.qtd;

          tsv += `${{item.cod}}\\t${{item.nome}}\\t${{item.unid}}\\t${{item.qtd.toFixed(2).replace('.', ',')}}\\t${{suppName}}\\t${{p.toFixed(4).replace('.', ',')}}\\t${{tot.toFixed(2).replace('.', ',')}}\\n`;
        }}
      }});

      navigator.clipboard.writeText(tsv).then(() => {{
        const btn = document.getElementById("btn-copy-sienge");
        if (btn) {{
          btn.innerHTML = `<svg class="w-4 h-4 text-emerald-400" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12.75l6 6 9-13.5"/></svg><span>Copiado com Sucesso!</span>`;
          setTimeout(() => {{
            btn.innerHTML = `<svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M15.75 17.25v3.375c0 .621-.504 1.125-1.125 1.125h-9.75a1.125 1.125 0 01-1.125-1.125V7.875c0-.621.504-1.125 1.125-1.125H6.75a9.06 9.06 0 011.5.124m7.5 10.376h3.375c.621 0 1.125-.504 1.125-1.125V11.25c0-4.46-3.243-8.161-7.5-8.876a9.06 9.06 0 00-1.5-.124H9.375c-.621 0-1.125.504-1.125 1.125v3.5m7.5 10.375H9.375a1.125 1.125 0 01-1.125-1.125v-9.25m12 6.625v-1.875a3.375 3.375 0 00-3.375-3.375h-1.5a1.125 1.125 0 01-1.125-1.125v-1.5a3.375 3.375 0 00-3.375-3.375H9.75"/></svg><span>Copiar TSV para Sienge</span>`;
          }}, 2500);
        }}
      }}).catch(() => {{
        alert("Pressione Ctrl+C para copiar.");
      }});
    }}

    // ==================== GRÁFICOS COMPARATIVOS (9 PROPOSTAS) ====================
    function renderBarCharts() {{
      const barsContainer = document.getElementById("chart-bars-container");
      const compContainer = document.getElementById("chart-comp-container");
      if (!barsContainer || !compContainer) return;

      barsContainer.innerHTML = "";
      compContainer.innerHTML = "";

      const calcS = (isCD, sIdx) => itemsData.reduce((acc, i) => {{
        const p = isCD ? i.pricesCD[sIdx] : i.pricesBR[sIdx];
        return acc + (p !== null ? p * i.qtd : 0);
      }}, 0);

      const suppliersAll = [
        {{ "nome": "ArcelorMittal (Barra Reta)", "val": calcS(false, 0), "color": "#1d4ed8" }},
        {{ "nome": "Gerdau (Barra Reta)", "val": calcS(false, 1), "color": "#2563eb" }},
        {{ "nome": "Paranafer (CD - S/ Arame)", "val": calcS(true, 0), "color": "#0891b2" }},
        {{ "nome": "Kauefer (Barra Reta)", "val": calcS(false, 2), "color": "#64748b" }},
        {{ "nome": "NORTEFER (Barra Reta)", "val": calcS(false, 3), "color": "#0284c7" }},
        {{ "nome": "ArcelorMittal (Corte e Dobra - 1º Compl.)", "val": calcS(true, 1), "color": "#0f766e" }},
        {{ "nome": "Gerdau (Corte e Dobra)", "val": calcS(true, 2), "color": "#3b82f6" }},
        {{ "nome": "NORTEFER (Corte e Dobra)", "val": calcS(true, 4), "color": "#475569" }},
        {{ "nome": "Kauefer (Corte e Dobra)", "val": calcS(true, 3), "color": "#334155" }}
      ];

      suppliersAll.sort((a, b) => a.val - b.val);
      const maxVal = Math.max(...suppliersAll.map(s => s.val));

      suppliersAll.forEach(s => {{
        const pct = maxVal > 0 ? (s.val / maxVal) * 100 : 0;
        const row = document.createElement("div");
        row.className = "space-y-1";
        row.innerHTML = `
          <div class="flex justify-between text-xs font-bold">
            <span class="text-slate-800">${{s.nome}}</span>
            <span class="text-slate-900 num-tabular">${{formatBRL(s.val)}}</span>
          </div>
          <div class="h-3 w-full bg-slate-100 rounded-full overflow-hidden">
            <div class="h-full rounded-full transition-all" style="width: ${{pct}}%; background-color: ${{s.color}}"></div>
          </div>
        `;
        barsContainer.appendChild(row);
      }});

      // Distribuição por bitola
      const totalVol = itemsData.reduce((acc, i) => acc + i.qtd, 0);
      itemsData.filter(i => i.qtd > 0).forEach(item => {{
        const pct = totalVol > 0 ? (item.qtd / totalVol) * 100 : 0;
        const row = document.createElement("div");
        row.className = "space-y-1";
        row.innerHTML = `
          <div class="flex justify-between text-xs">
            <span class="font-bold text-slate-800">${{item.nome}}</span>
            <span class="text-slate-500 font-semibold num-tabular">${{formatNum(item.qtd, 2)}} kg (${{pct.toFixed(1)}}%)</span>
          </div>
          <div class="h-2.5 w-full bg-slate-100 rounded-full overflow-hidden">
            <div class="h-full bg-[#0a2540] rounded-full transition-all" style="width: ${{pct}}%"></div>
          </div>
        `;
        compContainer.appendChild(row);
      }});
    }}

    // ==================== EXPORTAÇÃO E PERSISTÊNCIA ====================
    function toggleExportMenu() {{
      const el = document.getElementById("export-dropdown");
      if (el) el.classList.toggle("hidden");
    }}

    document.addEventListener("click", function(e) {{
      const menu = document.getElementById("export-dropdown");
      const btn = e.target.closest("button");
      if (menu && !menu.classList.contains("hidden")) {{
        if (!menu.contains(e.target) && (!btn || !btn.getAttribute("onclick") || !btn.getAttribute("onclick").includes("toggleExportMenu"))) {{
          menu.classList.add("hidden");
        }}
      }}
    }});

    function saveScenario() {{
      const state = {{
        selectedCD: selectedCD,
        selectedBR: selectedBR,
        quantities: itemsData.map(i => ({{ id: i.id, qtd: i.qtd }})),
        timestamp: new Date().toISOString()
      }};
      try {{
        localStorage.setItem("giardini_aco_cenario_v5", JSON.stringify(state));
        alert("Cenário salvo com sucesso no navegador!");
      }} catch (e) {{
        alert("Cenário salvo!");
      }}
    }}

    function exportToCSV() {{
      let csv = "Código Sienge;Descrição;Unidade;Volume (kg);Paranafer (CD);Arcelor (CD);Gerdau (CD);Kauefer (CD);NORTEFER (CD);Arcelor (Barra);Gerdau (Barra);Kauefer (Barra);NORTEFER (Barra)\\n";
      itemsData.forEach(i => {{
        csv += `${{i.cod}};"${{i.nome}}";${{i.unid}};${{i.qtd.toFixed(2).replace('.', ',')}};`;
        const prCD = i.pricesCD.map(p => p !== null ? p.toFixed(4).replace('.', ',') : "-").join(";");
        const prBR = i.pricesBR.map(p => p !== null ? p.toFixed(4).replace('.', ',') : "-").join(";");
        csv += `${{prCD}};${{prBR}}\\n`;
      }});

      const blob = new Blob(["\\uFEFF" + csv], {{ type: "text/csv;charset=utf-8;" }});
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "MAPA_COTACAO_GIARDINI_ACO_REV03.csv";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }}

    function exportToXLSX() {{
      let totalVolume = itemsData.reduce((acc, i) => acc + i.qtd, 0);
      let xml = `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Worksheet ss:Name="Corte e Dobra">
  <Table>
   <Row>
    <Cell><Data ss:Type="String">Código Sienge</Data></Cell>
    <Cell><Data ss:Type="String">Descrição do Insumo</Data></Cell>
    <Cell><Data ss:Type="String">Unidade</Data></Cell>
    <Cell><Data ss:Type="String">Quantidade (kg)</Data></Cell>
    <Cell><Data ss:Type="String">Paranafer (CD)</Data></Cell>
    <Cell><Data ss:Type="String">ArcelorMittal (CD)</Data></Cell>
    <Cell><Data ss:Type="String">Gerdau (CD)</Data></Cell>
    <Cell><Data ss:Type="String">Kauefer (CD)</Data></Cell>
    <Cell><Data ss:Type="String">NORTEFER (CD)</Data></Cell>
    <Cell><Data ss:Type="String">Total Selecionado (R$)</Data></Cell>
   </Row>`;

      itemsData.forEach(i => {{
        const sel = selectedCD[i.id];
        const pSel = (sel !== undefined && i.pricesCD[sel] !== null) ? i.pricesCD[sel] : 0;
        const tot = pSel * i.qtd;
        xml += `
   <Row>
    <Cell><Data ss:Type="String">${{i.cod}}</Data></Cell>
    <Cell><Data ss:Type="String">${{i.nome}}</Data></Cell>
    <Cell><Data ss:Type="String">${{i.unid}}</Data></Cell>
    <Cell><Data ss:Type="Number">${{i.qtd}}</Data></Cell>
    <Cell><Data ss:Type="${{i.pricesCD[0] ? "Number" : "String"}}">${{i.pricesCD[0] || "-"}}</Data></Cell>
    <Cell><Data ss:Type="${{i.pricesCD[1] ? "Number" : "String"}}">${{i.pricesCD[1] || "-"}}</Data></Cell>
    <Cell><Data ss:Type="${{i.pricesCD[2] ? "Number" : "String"}}">${{i.pricesCD[2] || "-"}}</Data></Cell>
    <Cell><Data ss:Type="${{i.pricesCD[3] ? "Number" : "String"}}">${{i.pricesCD[3] || "-"}}</Data></Cell>
    <Cell><Data ss:Type="${{i.pricesCD[4] ? "Number" : "String"}}">${{i.pricesCD[4] || "-"}}</Data></Cell>
    <Cell><Data ss:Type="Number">${{tot}}</Data></Cell>
   </Row>`;
      }});

      xml += `
  </Table>
 </Worksheet>
 <Worksheet ss:Name="Barra Reta">
  <Table>
   <Row>
    <Cell><Data ss:Type="String">Código Sienge</Data></Cell>
    <Cell><Data ss:Type="String">Descrição do Insumo</Data></Cell>
    <Cell><Data ss:Type="String">Unidade</Data></Cell>
    <Cell><Data ss:Type="String">Quantidade (kg)</Data></Cell>
    <Cell><Data ss:Type="String">ArcelorMittal (Barra)</Data></Cell>
    <Cell><Data ss:Type="String">Gerdau (Barra)</Data></Cell>
    <Cell><Data ss:Type="String">Kauefer (Barra)</Data></Cell>
    <Cell><Data ss:Type="String">NORTEFER (Barra)</Data></Cell>
    <Cell><Data ss:Type="String">Total Selecionado (R$)</Data></Cell>
   </Row>`;

      itemsData.forEach(i => {{
        const sel = selectedBR[i.id];
        const pSel = (sel !== undefined && i.pricesBR[sel] !== null) ? i.pricesBR[sel] : 0;
        const tot = pSel * i.qtd;
        xml += `
   <Row>
    <Cell><Data ss:Type="String">${{i.cod}}</Data></Cell>
    <Cell><Data ss:Type="String">${{i.nome}}</Data></Cell>
    <Cell><Data ss:Type="String">${{i.unid}}</Data></Cell>
    <Cell><Data ss:Type="Number">${{i.qtd}}</Data></Cell>
    <Cell><Data ss:Type="${{i.pricesBR[0] ? "Number" : "String"}}">${{i.pricesBR[0] || "-"}}</Data></Cell>
    <Cell><Data ss:Type="${{i.pricesBR[1] ? "Number" : "String"}}">${{i.pricesBR[1] || "-"}}</Data></Cell>
    <Cell><Data ss:Type="${{i.pricesBR[2] ? "Number" : "String"}}">${{i.pricesBR[2] || "-"}}</Data></Cell>
    <Cell><Data ss:Type="${{i.pricesBR[3] ? "Number" : "String"}}">${{i.pricesBR[3] || "-"}}</Data></Cell>
    <Cell><Data ss:Type="Number">${{tot}}</Data></Cell>
   </Row>`;
      }});

      xml += `
  </Table>
 </Worksheet>
</Workbook>`;

      const blob = new Blob([xml], {{ type: "application/vnd.ms-excel;charset=utf-8;" }});
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `MAPA_COTACAO_GIARDINI_ACO_${{totalVolume.toFixed(0)}}KG_REV03.xls`;
      a.click();
    }}

    // INICIALIZAÇÃO COMPLETA
    window.addEventListener("DOMContentLoaded", () => {{
      renderTableCD();
      renderTableBR();
      renderAuditTable();
      renderSiengeTable();
      runSim();
      switchTab("corte-dobra");
    }});
  </script>
</body>
</html>
"""

if __name__ == "__main__":
    update_excel()
    html_content = build_html()

    with open(TARGET_FILE, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"Saved: {TARGET_FILE} ({len(html_content)} bytes)")

    with open(BRAIN_FILE, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"Saved: {BRAIN_FILE} ({len(html_content)} bytes)")
