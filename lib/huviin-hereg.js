import {
  Document,
  Packer,
  Table,
  TableRow,
  TableCell,
  Paragraph,
  TextRun,
  AlignmentType,
  WidthType,
  SectionType,
} from 'docx';

const s = (v) => (v !== null && v !== undefined ? String(v).trim() : '');

function textParas(text, opts = {}) {
  const lines = (text || '').split('\n');
  return lines.map(
    (line) =>
      new Paragraph({
        children: [
          new TextRun({
            text: line,
            bold: opts.bold || false,
            color: opts.color || undefined,
          }),
        ],
        alignment: opts.alignment || AlignmentType.LEFT,
      })
  );
}

function makeCell(text, colSpan = 1, rowSpan = 1, opts = {}) {
  return new TableCell({
    columnSpan: colSpan,
    rowSpan: rowSpan,
    children: textParas(text, opts),
  });
}

function buildRows(row) {
  const coord = (() => {
    let c = '1.Өргөрөг: ' + s(row[14]) + '\nУртраг: ' + s(row[15]);
    if (row[16]) c += '\n2.Өргөрөг: ' + s(row[16]) + '\nУртраг: ' + s(row[17]);
    return c;
  })();

  return [
    // 0: Газар зүйн нэр
    new TableRow({
      children: [
        makeCell('Газар зүйн нэр /монгол, латин галиг/', 4),
        makeCell(s(row[3]), 3),
        makeCell(s(row[4]), 3),
      ],
    }),
    // 1: Дахин давтагдашгүй дугаар
    new TableRow({
      children: [
        makeCell('Газар зүйн нэрийн дахин давтагдашгүй дугаар', 4),
        makeCell(s(row[1]), 6),
      ],
    }),
    // 2: Гарал, үүсэл
    new TableRow({
      children: [
        makeCell('Газар зүйн нэрийн гарал, үүсэл', 4),
        makeCell(s(row[13]) || 'Уламжлалт нэр', 6),
      ],
    }),
    // 3: Төрөл
    new TableRow({
      children: [
        makeCell('Газар зүйн нэрийн төрөл /дэвсгэр нэр/', 4),
        makeCell(s(row[5]), 6),
      ],
    }),
    // 4: Аймаг, сум, баг
    new TableRow({
      children: [
        makeCell('Харьяалагдах аймаг, сум, баг', 4),
        makeCell(s(row[18]), 6),
      ],
    }),
    // 5: Байрлал, тайлбар
    new TableRow({
      children: [
        makeCell('Газар зүйн нэрийн ерөнхий байрлал, тайлбар', 4),
        makeCell(s(row[19]), 6),
      ],
    }),
    // 6: Солбицол
    new TableRow({
      children: [
        makeCell('Газар зүйн нэрийн солбицол, UTM, 48-р бүс', 4),
        makeCell(coord, 6),
      ],
    }),
    // 7: Зургийн нэрлэвэр
    new TableRow({
      children: [
        makeCell(
          'Газар зүйн нэрийн орших 1:25 000-ны масштабтай байр зүйн зургийн нэрлэвэр',
          4
        ),
        makeCell(s(row[9]), 2),
        makeCell('Нэрийн зургийн индекс', 2),
        makeCell(s(row[2]), 2),
      ],
    }),
    // 8: Масштаб header (rowSpan=2 for left cell)
    new TableRow({
      children: [
        makeCell('', 4, 2),
        makeCell('1:25000 зурагт', 3),
        makeCell('1:100 000 зурагт', 3),
      ],
    }),
    // 9: Масштаб values
    new TableRow({
      children: [makeCell('', 3), makeCell(s(row[10]), 3)],
    }),
    // 10: Актын нэр (red)
    new TableRow({
      children: [
        makeCell('Газар зүйн нэрийг баталгаажуулсан актын нэр, дугаар, огноо', 4),
        makeCell(s(row[20]), 6, 1, { color: 'FF0000' }),
      ],
    }),
    // 11: Өөрчлөлт
    new TableRow({
      children: [
        makeCell('Өөрчлөлт орсон эсэх, шалтгаан', 4),
        makeCell('', 6),
      ],
    }),
    // 12: Байршлын зураг header
    new TableRow({
      children: [
        makeCell('Газар зүйн нэрийн байршлын зураг', 10, 1, {
          bold: true,
          alignment: AlignmentType.CENTER,
        }),
      ],
    }),
    // 13: Image placeholder
    new TableRow({
      children: [makeCell('', 4), makeCell('', 6)],
    }),
    // 14: Нотлох баримт
    new TableRow({
      children: [
        makeCell('Газар зүйн нэрийн тодруулалтын үеийн нотлох баримт:', 4),
        makeCell('Аудио, видео бичлэг: □ \nТэмдэглэл:    □        Фото зураг:   □', 6),
      ],
    }),
    // 15: Тодруулсан иргэн
    new TableRow({
      children: [
        makeCell('Газар зүйн нэрийг тодруулсан иргэн, хуулийн этгээд', 4),
        makeCell(
          '"Инженер геодези" ХХК-ны инженер:\nМУ-ын зөвлөх инженер Д.Оюунчимэг\nИнженер: Э.Ануун, Н.Бумчин',
          6
        ),
      ],
    }),
    // 16: Газарчин
    new TableRow({
      children: [
        makeCell('Газар зүйн нэрийг тодруулсан газарчин /орон нутгийн/', 4),
        makeCell('Н.Очирваань, багийн өндөр настан\nЭ.Эрдэнэтунгалаг, газрын даамал', 6),
      ],
    }),
    // 17: Огноо
    new TableRow({
      children: [
        makeCell('Газар зүйн нэрийн хувийн хэрэг хөтөлсөн:', 4),
        makeCell('/2024 оны 05-р сарын 15-ны өдөр/', 6),
      ],
    }),
  ];
}

export async function generateHuviinHereg(excelRows) {
  const sections = excelRows.map((row, i) => ({
    properties: {
      type: i === 0 ? SectionType.CONTINUOUS : SectionType.NEXT_PAGE,
    },
    children: [
      new Paragraph({
        children: [new TextRun({ text: 'Газар зүйн нэрийн хувийн хэрэг', bold: true, size: 26 })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 100 },
      }),
      new Table({
        rows: buildRows(row),
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: [900, 900, 900, 900, 900, 900, 900, 900, 900, 900],
      }),
    ],
  }));

  const doc = new Document({ sections });
  return Packer.toBuffer(doc);
}
