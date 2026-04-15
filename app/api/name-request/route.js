import { NextResponse } from 'next/server';
import { read, utils } from 'xlsx';
import { generateNameRequest } from '../../../lib/name-request-doc';

export async function POST(request) {
  try {
    const formData = await request.formData();
    const file = formData.get('file');
    if (!file) {
      return NextResponse.json({ error: 'Файл олдсонгүй' }, { status: 400 });
    }

    const buffer = Buffer.from(await file.arrayBuffer());
    const wb = read(buffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const allRows = utils.sheet_to_json(ws, { header: 1, defval: null });
    const dataRows = allRows.slice(2);

    if (dataRows.length === 0) {
      return NextResponse.json(
        { error: 'Файлд өгөгдөл олдсонгүй (3-р мөрөөс эхлэх ёстой).' },
        { status: 400 }
      );
    }

    const docBuffer = await generateNameRequest(dataRows);

    return new NextResponse(docBuffer, {
      status: 200,
      headers: {
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': 'attachment; filename="huseltiing_maygt.docx"',
      },
    });
  } catch (err) {
    console.error(err);
    return NextResponse.json({ error: 'Алдаа гарлаа: ' + err.message }, { status: 500 });
  }
}
