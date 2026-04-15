import { NextResponse } from 'next/server';
import { read, utils } from 'xlsx';
import { generateHuviinHereg } from '../../../lib/huviin-hereg';
import { filterAndValidate, HUVIIN_HEREG_COLUMNS } from '../../../lib/validate-rows';

export async function POST(request) {
  try {
    console.log('[huviin-hereg] POST request received');

    const formData = await request.formData();
    const file = formData.get('file');
    if (!file) {
      return NextResponse.json({ error: 'Файл олдсонгүй' }, { status: 400 });
    }

    console.log(`[huviin-hereg] File: ${file.name}, size: ${file.size} bytes`);

    const buffer = Buffer.from(await file.arrayBuffer());
    const wb = read(buffer, { type: 'buffer' });
    console.log(`[huviin-hereg] Sheets: ${wb.SheetNames}`);

    const ws = wb.Sheets[wb.SheetNames[0]];
    const allRows = utils.sheet_to_json(ws, { header: 1, defval: null });
    console.log(`[huviin-hereg] Excel нийт мөр: ${allRows.length}`);

    const rawRows = allRows.slice(2); // header 2 мөр хасна
    const result = filterAndValidate(rawRows, HUVIIN_HEREG_COLUMNS);

    if (result.error) {
      return NextResponse.json({ error: result.error }, { status: 400 });
    }

    console.log(`[huviin-hereg] Generating Word document for ${result.dataRows.length} rows...`);
    const docBuffer = await generateHuviinHereg(result.dataRows);
    console.log(`[huviin-hereg] Done, size: ${docBuffer.length} bytes`);

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
