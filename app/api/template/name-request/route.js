import { NextResponse } from 'next/server';
import { utils, write } from 'xlsx';

export async function GET() {
  const wb = utils.book_new();

  const rows = [
    // Row 1: Гарчиг
    ['Газар зүйн нэрийг шинээр өгөх хүсэлтийн жагсаалт'],
    // Row 2: Баганын гарчиг
    [
      'A', 'B', 'C',
      'D - Газар зүйн нэр *',
      'E',
      'F - Дэвсгэр нэр',
      'G', 'H', 'I', 'J', 'K', 'L', 'M',
      'N - Гарал үүсэл',
      'O - Өргөрөг 1 *',
      'P - Уртраг 1 *',
      'Q - Өргөрөг 2',
      'R - Уртраг 2',
      'S - Аймаг/сум/баг *',
      'T - Алслагдах зай / байрлал',
    ],
    // Row 3: Жишээ өгөгдөл
    [
      '', '', '',
      'Хайрхан уул',
      '',
      'уул',
      '', '', '', '', '', '', '',
      'Уламжлалт нэр',
      '43°30\'15"',
      '104°12\'30"',
      '', '',
      'Өмнөговь аймаг, Гурвантэс сум, 1-р баг',
      'Сумын төвөөс 25 км зүүн хойш',
    ],
  ];

  const ws = utils.aoa_to_sheet(rows);
  ws['!cols'] = Array(20).fill({ wch: 22 });
  ws['!cols'][0] = { wch: 4 };

  utils.book_append_sheet(wb, ws, 'Sheet1');
  const buffer = write(wb, { type: 'buffer', bookType: 'xlsx' });

  return new NextResponse(buffer, {
    status: 200,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="name_request_template.xlsx"',
    },
  });
}
