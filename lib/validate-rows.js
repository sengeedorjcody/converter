// Хувийн хэрэгт шаардлагатай баганууд: { index: [нэр, заавал эсэх] }
const HUVIIN_HEREG_COLUMNS = {
  1:  ['Дахин давтагдашгүй дугаар', true],
  2:  ['Нэрийн зургийн индекс',     true],
  3:  ['Газар зүйн нэр',            true],
  4:  ['Төрөл',                     true],
  5:  ['Дэвсгэр нэр / Ангилал',     false],
  9:  ['Байр зүйн зургийн нэрлэвэр',false],
  10: ['1:100 000 зурагт',          false],
  13: ['Гарал үүсэл',               false],
  14: ['Өргөрөг 1',                 true],
  15: ['Уртраг 1',                  true],
  16: ['Өргөрөг 2',                 false],
  17: ['Уртраг 2',                  false],
  18: ['Аймаг/сум/баг',             true],
  19: ['Байрлал тайлбар',           false],
  20: ['Актын дугаар',              false],
};

const NAME_REQUEST_COLUMNS = {
  3:  ['Газар зүйн нэр',  true],
  5:  ['Дэвсгэр нэр',     false],
  14: ['Өргөрөг 1',       true],
  15: ['Уртраг 1',        true],
  16: ['Өргөрөг 2',       false],
  17: ['Уртраг 2',        false],
  18: ['Аймаг/сум/баг',   true],
  19: ['Байрлал тайлбар', false],
};

const val = (row, idx) => {
  const v = row[idx];
  return v !== null && v !== undefined && String(v).trim() !== '';
};

/**
 * Excel-ийн raw мөрүүдийг (header хасагдсан) шүүж, validate хийнэ.
 * col[3] (Газар зүйн нэр) хоосон болмогц зогсоно.
 * Алдаа байвал { error: string } буцаана.
 * Амжилттай бол { dataRows: array } буцаана — docx-д шууд дамжуулах боломжтой.
 */
export function filterAndValidate(rawRows, columns = HUVIIN_HEREG_COLUMNS) {
  const dataRows = [];

  for (let i = 0; i < rawRows.length; i++) {
    const row = rawRows[i];
    const excelRowNum = i + 3; // header 2 мөр байна
    const nameVal = row[3];
    const isEmpty = nameVal === null || nameVal === undefined || String(nameVal).trim() === '';

    if (isEmpty) {
      if (dataRows.length > 0) {
        console.log(`[validate] Excel мөр ${excelRowNum} — хоосон мөр, боловсруулалт зогслоо.`);
        break;
      } else {
        console.log(`[validate] Excel мөр ${excelRowNum} — өмнөх хоосон мөр, алгасав.`);
        continue;
      }
    }

    // Утга бүхий мөрийн log
    console.log(`[validate] Excel мөр ${excelRowNum}:`);
    row.forEach((v, ci) => {
      if (v !== null && v !== undefined && String(v).trim()) {
        console.log(`  col[${ci}] = ${String(v).slice(0, 80)}`);
      }
    });

    dataRows.push({ excelRowNum, row });
  }

  console.log(`[validate] Боловсруулах мөр: ${dataRows.length}`);

  if (dataRows.length === 0) {
    return { error: 'Боловсруулах өгөгдөл олдсонгүй. D баганад (Газар зүйн нэр) утга байхгүй байна.' };
  }

  // Validation
  const errors = [];
  for (const { excelRowNum, row } of dataRows) {
    const missing = [];
    for (const [idx, [name, required]] of Object.entries(columns)) {
      if (required && !val(row, Number(idx))) {
        missing.push(name);
        console.warn(`  [Мөр ${excelRowNum}] ❌ ЗААВАЛ хоосон: col[${idx}] = '${name}'`);
      }
    }
    if (missing.length > 0) {
      errors.push(`Excel мөр ${excelRowNum}: ${missing.join(', ')}`);
    }
  }

  if (errors.length > 0) {
    console.warn(`[validate] ${errors.length} мөрт алдаа байна`);
    return { error: 'Дараах мөрүүдэд заавал талбар дутуу байна:\n' + errors.join('\n') };
  }

  return { dataRows: dataRows.map(({ row }) => row) };
}

export { HUVIIN_HEREG_COLUMNS, NAME_REQUEST_COLUMNS };
