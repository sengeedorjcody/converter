import { NextResponse } from 'next/server';
import { read, utils } from 'xlsx';
import { generateHuviinHereg } from '../../../lib/huviin-hereg';

export async function POST(request) {
  try {
    console.log('[huviin-hereg] POST request received');

    const formData = await request.formData();
    const file = formData.get('file');
    if (!file) {
      console.warn('[huviin-hereg] No file in formData');
      return NextResponse.json({ error: 'Файл олдсонгүй' }, { status: 400 });
    }

    console.log(`[huviin-hereg] File: ${file.name}, size: ${file.size} bytes`);

    const buffer = Buffer.from(await file.arrayBuffer());
    console.log(`[huviin-hereg] Buffer size: ${buffer.length}`);

    const wb = read(buffer, { type: 'buffer' });
    console.log(`[huviin-hereg] Sheets: ${wb.SheetNames}`);

    const ws = wb.Sheets[wb.SheetNames[0]];
    const allRows = utils.sheet_to_json(ws, { header: 1, defval: null });
    console.log(`[huviin-hereg] Total rows (incl. headers): ${allRows.length}`);

    const dataRows = allRows.slice(2);
    console.log(`[huviin-hereg] Data rows: ${dataRows.length}`);

    if (dataRows.length === 0) {
      console.warn('[huviin-hereg] No data rows found');
      return NextResponse.json(
        { error: 'Файлд өгөгдөл олдсонгүй (3-р мөрөөс эхлэх ёстой).' },
        { status: 400 }
      );
    }

    // Log first row column values for debugging
    const firstRow = dataRows[0];
    console.log('[huviin-hereg] First row column count:', firstRow.length);
    firstRow.forEach((val, i) => {
      if (val !== null && val !== undefined) {
        console.log(`  col[${i}] = ${String(val).slice(0, 60)}`);
      }
    });

    console.log('[huviin-hereg] Generating Word document...');
    const docBuffer = await generateHuviinHereg(dataRows);
    console.log(`[huviin-hereg] Document generated, size: ${docBuffer.length} bytes`);

    return new NextResponse(docBuffer, {
      status: 200,
      headers: {
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': 'attachment; filename="Huviin_hereg.docx"',
      },
    });
  } catch (err) {
    console.error('[huviin-hereg] ERROR:', err);
    return NextResponse.json({ error: 'Алдаа гарлаа: ' + err.message }, { status: 500 });
  }
}
