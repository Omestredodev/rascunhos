Você é um especialista em Microsoft Excel Online + Office Scripts (TypeScript), especialmente em ambientes corporativos com regras restritivas do compilador (erros do tipo “Office Scripts cannot infer the data type of this variable. Please declare a type for the variable” e proibição de any).

Eu preciso de um Office Script que leia uma base no Excel e gere um relatório em árvore/cascata de 2 níveis para comparar Notificações vs Advertências, por ano, com regras específicas.

Estrutura da base
Período: 2018 a 2023
Cabeçalho começa na linha 4
A linha 5 deve ser ignorada totalmente (não entra em contagem nem em nada)
Dados efetivos começam na linha 6
A coluna C tem o campo "Data" (usar para obter o ano)
A coluna L é "Referência" (texto do título)
Regras de identificação do tipo (Notificação/Advertência)
Considerar uma linha se, em Referência (L), existir ocorrência de:
“notificação” ou “advertência”
Deve ser:
case-insensitive
tolerante a ausência de acentos/cedilha, ou seja:
“notificacao” também conta (sem “ç”/acentos)
“advertencia” também conta
ignorar caractere ^ caso apareça no texto (ex.: “notifica^cao”)
Regras de cancelamento / exclusão
Coluna W tem cabeçalho "CANCELADO"
se tiver "x" ou "X", a linha deve ser ignorada
Além disso: se a linha for notificação/advertência, verificar a formatação da célula L (Referência):
se a cor da fonte estiver em qualquer tom de vermelho (não só #FF0000), a linha também deve ser ignorada
Implementar detecção robusta de “vermelho” aceitando #RRGGBB e #AARRGGBB, usando threshold (ex.: R alto e G/B baixos)
Regras de área/classificação (colunas O/P/Q/R)
Colunas:
O = Vendas
P = Serviços & Peças
Q = StarClass
R = Outros
Se a célula estiver preenchida com qualquer coisa, considerar como “marcado”
Se mais de uma dessas colunas estiver preenchida na mesma linha:
não incluir na soma
registrar o número da linha e listar no final do relatório
Caso especial Vans:
Se a única coluna marcada (entre O/P/Q/R) for preenchida com exatamente “Vans” ou “VANS” (sem nada além):
colocar em um relatório à parte chamado VANS ONLY
no final, gerar dois blocos: COM Vans e SEM Vans
Saída desejada (relatório)
Gerar/atualizar uma aba de saída chamada:

Relatorio_Notif_Advert_2N
Nela, imprimir três blocos, cada um em formato 2 níveis:

SEM VANS
COM VANS
VANS ONLY
Formato em cascata (2 níveis):

Para cada ano de 2018 a 2023:
Linha do Ano com total anual
Linha “Notificação” com contagem por área (O/P/Q/R) e total do tipo
Linha “Advertência” com contagem por área e total do tipo
Linha em branco para separar anos (opcional)
Ao final: listar as linhas com múltiplas marcações em O/P/Q/R
Restrições do ambiente (muito importante)
O Office Scripts no meu tenant corporativo:
não permite any
frequentemente falha em inferência de tipos, especialmente em acessos como values[i][col]
pede tipos explícitos para variáveis e arrays
Portanto, escreva o script de modo “corporate-proof”, por exemplo:
usar const rowValues: ExcelScript.CellValue[] = values[i]; antes de indexar colunas
evitar Map, Record, reduce, map se isso causar inferência
pode usar matriz pré-alocada (ex.: 6 anos × 2 tipos × 4 áreas)
Entrega
Forneça o script completo pronto para colar em Automate > New Script
Explique rapidamente onde o relatório será gerado e como ele lida com:
linha 5 ignorada
cancelamento por W=X
cancelamento por fonte vermelha
multi-marcação O/P/Q/R
VANS ONLY / COM VANS / SEM VANS

Atualmente meu código é esse:
type TipoIdx = 0 | 1;            // 0=Notificação, 1=Advertência
type AreaIdx = 0 | 1 | 2 | 3;    // 0=Vendas, 1=Serviços & Peças, 2=StarClass, 3=Outros

function main(workbook: ExcelScript.Workbook) {
  const ws: ExcelScript.Worksheet = workbook.getActiveWorksheet();

  // ===== CONFIG =====
  const SKIP_ROW: number = 5;        // ignorar totalmente
  const FIRST_DATA_ROW: number = 6;  // dados começam aqui

  const DATE_COL: number = 3;        // C
  const REF_COL: number = 12;        // L
  const SALES_COL: number = 15;      // O
  const SERVICES_COL: number = 16;   // P
  const STAR_COL: number = 17;       // Q
  const OTHER_COL: number = 18;      // R
  const CANCEL_COL: number = 23;     // W

  const YEAR_START: number = 2018;
  const YEAR_END: number = 2023;

  // ===== RANGE / VALUES =====
  const used: ExcelScript.Range | undefined = ws.getUsedRange();
  if (!used) throw new Error("A planilha não tem dados (usedRange vazio).");

  const lastRow1: number = used.getRowIndex() + used.getRowCount(); // 1-based
  if (lastRow1 < FIRST_DATA_ROW) throw new Error(`Não há dados a partir da linha ${FIRST_DATA_ROW}.`);

  // Lê A..W (23 colunas) desde linha 6
  const dataRange: ExcelScript.Range = ws.getRangeByIndexes(
    FIRST_DATA_ROW - 1, // 0-based
    0,
    lastRow1 - FIRST_DATA_ROW + 1,
    CANCEL_COL
  );
  const values: ExcelScript.CellValue[][] = dataRange.getValues();

  // ===== CONTADORES [6][2][4] =====
  const semVans: number[][][] = cube6x2x4();
  const comVans: number[][][] = cube6x2x4();
  const vansOnly: number[][][] = cube6x2x4();

  const multiFillRows: number[] = [];

  // ===== LOOP =====
  for (let i: number = 0; i < values.length; i++) {
    const row1: number = FIRST_DATA_ROW + i;
    if (row1 === SKIP_ROW) continue;

    // Truque corporativo: tipa a linha antes de indexar colunas
    const rowValues: ExcelScript.CellValue[] = values[i];

    const dateVal: ExcelScript.CellValue = rowValues[DATE_COL - 1];
    const refVal: ExcelScript.CellValue = rowValues[REF_COL - 1];
    const cancelVal: ExcelScript.CellValue = rowValues[CANCEL_COL - 1];

    // CANCELADO por X?
    if (isCancelledByX(cancelVal)) continue;

    // Tipo (Notificação vs Advertência)
    const tipoIdx: TipoIdx | null = getTipoIdx(refVal);
    if (tipoIdx === null) continue;

    // CANCELADO por fonte vermelha em L?
    const refCell: ExcelScript.Range = ws.getCell(row1 - 1, REF_COL - 1);
    const fontColor: string = refCell.getFormat().getFont().getColor();
    if (isRedFontByThreshold(fontColor)) continue;

    // Ano pela Data (C)
    const ano: number = getYear(dateVal);
    if (ano < YEAR_START || ano > YEAR_END) continue;
    const anoIdx: number = ano - YEAR_START; // 0..5

    // Áreas O/P/Q/R
    const o: ExcelScript.CellValue = rowValues[SALES_COL - 1];
    const p: ExcelScript.CellValue = rowValues[SERVICES_COL - 1];
    const q: ExcelScript.CellValue = rowValues[STAR_COL - 1];
    const r: ExcelScript.CellValue = rowValues[OTHER_COL - 1];

    const cls: AreaClass = classifyArea(o, p, q, r);

    if (cls.filledCount >= 2) {
      multiFillRows.push(row1);
      continue;
    }
    if (cls.filledCount === 0) continue;

    // Soma COM VANS
    comVans[anoIdx][tipoIdx][cls.areaIdx]++;

    // Soma SEM VANS / VANS ONLY
    if (cls.isVansOnly) vansOnly[anoIdx][tipoIdx][cls.areaIdx]++;
    else semVans[anoIdx][tipoIdx][cls.areaIdx]++;
  }

  // ===== OUTPUT =====
  const outName: string = "Relatorio_Notif_Advert_2N";
  let outWs: ExcelScript.Worksheet;
  const existing: ExcelScript.Worksheet | undefined = workbook.getWorksheet(outName);
  if (existing) outWs = existing;
  else outWs = workbook.addWorksheet(outName);

  const outUsed: ExcelScript.Range | undefined = outWs.getUsedRange();
  if (outUsed) outUsed.clear(ExcelScript.ClearApplyTo.all);

  let outRow: number = 1;
  outWs.getRange(`A${outRow}`).setValue("Relatório anual (2 níveis): Ano -> Tipo (Notificação/Advertência) -> Área");
  outRow++;
  outWs.getRange(`A${outRow}`).setValue("Regras: ignora linha 5; W=X ignora; Referência(L) vermelha ignora; multi-área não soma; VANS-only separado; 2018–2023.");
  outRow += 2;

  outRow = printBlock(outWs, outRow, "1) SEM VANS", semVans, YEAR_START, YEAR_END);
  outRow += 2;
  outRow = printBlock(outWs, outRow, "2) COM VANS", comVans, YEAR_START, YEAR_END);
  outRow += 2;
  outRow = printBlock(outWs, outRow, "3) VANS ONLY (única marcação é exatamente 'VANS')", vansOnly, YEAR_START, YEAR_END);
  outRow += 2;

  outWs.getRange(`A${outRow}`).setValue("Linhas com mais de uma coluna preenchida entre O/P/Q/R (não somadas):");
  outRow++;

  if (multiFillRows.length === 0) {
    outWs.getRange(`A${outRow}`).setValue("(nenhuma)");
    outRow++;
  } else {
    for (let k: number = 0; k < multiFillRows.length; k++) {
      outWs.getRange(`A${outRow}`).setValue(multiFillRows[k]);
      outRow++;
    }
  }

  const finalUsed: ExcelScript.Range | undefined = outWs.getUsedRange();
  if (finalUsed) finalUsed.getFormat().autofitColumns();
}

/* =========================
   Tipos auxiliares simples
========================= */

type AreaClass = { filledCount: number; areaIdx: AreaIdx; isVansOnly: boolean };

/* =========================
   Cubo fixo [6][2][4]
========================= */

function cube6x2x4(): number[][][] {
  // 6 anos (2018..2023), 2 tipos, 4 áreas
  const c: number[][][] = [
    [[0, 0, 0, 0], [0, 0, 0, 0]],
    [[0, 0, 0, 0], [0, 0, 0, 0]],
    [[0, 0, 0, 0], [0, 0, 0, 0]],
    [[0, 0, 0, 0], [0, 0, 0, 0]],
    [[0, 0, 0, 0], [0, 0, 0, 0]],
    [[0, 0, 0, 0], [0, 0, 0, 0]]
  ];
  return c;
}

/* =========================
   Impressão do bloco (2 níveis)
========================= */

function printBlock(
  ws: ExcelScript.Worksheet,
  startRow: number,
  title: string,
  cube: number[][][],
  yearStart: number,
  yearEnd: number
): number {
  let r: number = startRow;

  ws.getRange(`A${r}`).setValue(title);
  r++;

  const header: (string | number)[] = ["Ano", "Tipo", "Vendas", "Serviços & Peças", "StarClass", "Outros", "Total Tipo", "Total Ano"];
  ws.getRangeByIndexes(r - 1, 0, 1, 8).setValues([header]);
  r++;

  for (let ano: number = yearStart; ano <= yearEnd; ano++) {
    const idx: number = ano - yearStart;

    const notif: number[] = cube[idx][0];
    const advert: number[] = cube[idx][1];

    const notifTotal: number = notif[0] + notif[1] + notif[2] + notif[3];
    const advertTotal: number = advert[0] + advert[1] + advert[2] + advert[3];
    const totalAno: number = notifTotal + advertTotal;

    // Monta linhas em variáveis tipadas (evita inferência ruim no setValues)
    const rowAno: (string | number)[] = [ano, "", "", "", "", "", "", totalAno];
    ws.getRangeByIndexes(r - 1, 0, 1, 8).setValues([rowAno]);
    r++;

    const rowNotif: (string | number)[] = ["", "Notificação", notif[0], notif[1], notif[2], notif[3], notifTotal, ""];
    ws.getRangeByIndexes(r - 1, 0, 1, 8).setValues([rowNotif]);
    r++;

    const rowAdv: (string | number)[] = ["", "Advertência", advert[0], advert[1], advert[2], advert[3], advertTotal, ""];
    ws.getRangeByIndexes(r - 1, 0, 1, 8).setValues([rowAdv]);
    r++;

    r++; // linha em branco
  }

  return r;
}

/* =========================
   Regras / parsing
========================= */

function cellToString(v: ExcelScript.CellValue): string {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

function normalizeText(v: ExcelScript.CellValue): string {
  const s: string = cellToString(v).toLowerCase();
  return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\^/g, "");
}

function getTipoIdx(refVal: ExcelScript.CellValue): TipoIdx | null {
  const t: string = normalizeText(refVal);
  if (t.indexOf("notific") >= 0) return 0;
  if (t.indexOf("advert") >= 0) return 1;
  return null;
}

function isCancelledByX(v: ExcelScript.CellValue): boolean {
  const s: string = cellToString(v);
  return s.toUpperCase() === "X";
}

function getYear(v: ExcelScript.CellValue): number {
  if (typeof v === "number") {
    // Excel serial date
    const ms: number = Math.round((v - 25569) * 86400 * 1000);
    const d: Date = new Date(ms);
    return isNaN(d.getTime()) ? 0 : d.getFullYear();
  }

  const s: string = cellToString(v);
  if (!s) return 0;

  const d: Date = new Date(s);
  if (!isNaN(d.getTime())) return d.getFullYear();

  return 0;
}

function classifyArea(
  o: ExcelScript.CellValue,
  p: ExcelScript.CellValue,
  q: ExcelScript.CellValue,
  r: ExcelScript.CellValue
): AreaClass {
  const so: string = cellToString(o);
  const sp: string = cellToString(p);
  const sq: string = cellToString(q);
  const sr: string = cellToString(r);

  const hasO: boolean = so !== "";
  const hasP: boolean = sp !== "";
  const hasQ: boolean = sq !== "";
  const hasR: boolean = sr !== "";

  const filledCount: number = (hasO ? 1 : 0) + (hasP ? 1 : 0) + (hasQ ? 1 : 0) + (hasR ? 1 : 0);

  let areaIdx: AreaIdx = 0;
  let isVansOnly: boolean = false;

  if (filledCount === 1) {
    if (hasO) { areaIdx = 0; isVansOnly = (so.toUpperCase() === "VANS"); }
    else if (hasP) { areaIdx = 1; isVansOnly = (sp.toUpperCase() === "VANS"); }
    else if (hasQ) { areaIdx = 2; isVansOnly = (sq.toUpperCase() === "VANS"); }
    else { areaIdx = 3; isVansOnly = (sr.toUpperCase() === "VANS"); }
  }

  const result: AreaClass = { filledCount: filledCount, areaIdx: areaIdx, isVansOnly: isVansOnly };
  return result;
}

/**
 * Detecta fonte vermelha com tolerância (variações):
 * - aceita #RRGGBB e #AARRGGBB
 * - considera vermelho se R alto e G/B baixos
 */
function isRedFontByThreshold(color: string): boolean {
  const raw: string = (color ?? "").trim().toLowerCase();
  if (!raw) return false;
  if (raw === "red") return true;

  let hex: string = raw.startsWith("#") ? raw.substring(1) : raw;

  // ARGB: AARRGGBB -> RRGGBB
  if (hex.length === 8) hex = hex.substring(2);
  if (hex.length !== 6) return false;

  const r: number = parseInt(hex.substring(0, 2), 16);
  const g: number = parseInt(hex.substring(2, 4), 16);
  const b: number = parseInt(hex.substring(4, 6), 16);

  if (isNaN(r) || isNaN(g) || isNaN(b)) return false;

  // Threshold ajustável
  return (r >= 160 && g <= 120 && b <= 120);
}
