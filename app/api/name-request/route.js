import { NextResponse } from 'next/server';
import { read, utils } from 'xlsx';
import { generateNameRequest } from '../../../lib/name-request-doc';
import { filterAndValidate, NAME_REQUEST_COLUMNS } from '../../../lib/validate-rows';

export async function POST(request) {
  try {
    console.log('[name-request] POST request received');

    const formData = await request.formData();
    const file = formData.get('file');
    if (!file) {
      return NextResponse.json({ error: 'Файл олдсонгүй' }, { status: 400 });
    }

    console.log(`[name-request] File: ${file.name}, size: ${file.size} bytes`);

    const buffer = Buffer.from(await file.arrayBuffer());
    const wb = read(buffer, { type: 'buffer' });
    console.log(`[name-request] Sheets: ${wb.SheetNames}`);

    const ws = wb.Sheets[wb.SheetNames[0]];
    const allRows = utils.sheet_to_json(ws, { header: 1, defval: null });
    console.log(`[name-request] Excel нийт мөр: ${allRows.length}`);

    const rawRows = allRows.slice(2);
    const result = filterAndValidate(rawRows, NAME_REQUEST_COLUMNS);

    if (result.error) {
      return NextResponse.json({ error: result.error }, { status: 400 });
    }

    console.log(`[name-request] Generating Word document for ${result.dataRows.length} rows...`);
    const docBuffer = await generateNameRequest(result.dataRows);
    console.log(`[name-request] Done, size: ${docBuffer.length} bytes`);

    return new NextResponse(docBuffer, {
      status: 200,
      headers: {
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': 'attachment; filename="huseltiing_maygt.docx"',
      },
    });
  } catch (err) {
    console.error('[name-request] ERROR:', err);
    return NextResponse.json({ error: 'Алдаа гарлаа: ' + err.message }, { status: 500 });
  }
}
