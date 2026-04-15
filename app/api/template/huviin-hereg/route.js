import { NextResponse } from 'next/server';
import { utils, write } from 'xlsx';

export async function GET() {
  const wb = utils.book_new();

  const rows = [
    // Row 1: Гарчиг
    ['Газар зүйн нэрийн жагсаалт'],
    // Row 2: Баганын гарчиг (0-based: A=0, B=1, ...)
    [
      'A',
      'B - Дахин давтагдашгүй дугаар *',
      'C - Нэрийн зургийн индекс *',
      'D - Газар зүйн нэр *',
      'E - Төрөл *',
      'F - Дэвсгэр нэр / Ангилал',
      'G', 'H', 'I',
      'J - Байр зүйн зургийн нэрлэвэр',
      'K - 1:100 000 зурагт',
      'L', 'M',
      'N - Гарал үүсэл',
      'O - Өргөрөг 1 *',
      'P - Уртраг 1 *',
      'Q - Өргөрөг 2',
      'R - Уртраг 2',
      'S - Аймаг/сум/баг *',
      'T - Байрлал тайлбар',
      'U - Актын дугаар',
    ],
    // Row 3: Жишээ өгөгдөл
    [
      '', 'УГ-001', '42', 'Хайрхан уул', 'Уул', 'уул',
      '', '', '',
      'K-48-14', '14-42',
      '', '',
      'Уламжлалт нэр',
      '43°30\'15"', '104°12\'30"', '', '',
      'Өмнөговь аймаг, Гурвантэс сум, 1-р баг',
      'Сумын төвөөс 25 км зүүн хойш',
      '',
    ],
  ];

  const ws = utils.aoa_to_sheet(rows);

  // Баганын өргөн тохируулах
  ws['!cols'] = Array(21).fill({ wch: 22 });
  ws['!cols'][0] = { wch: 4 };

  utils.book_append_sheet(wb, ws, 'Sheet1');
  const buffer = write(wb, { type: 'buffer', bookType: 'xlsx' });

  return new NextResponse(buffer, {
    status: 200,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="huviin_hereg_template.xlsx"',
    },
  });
}
