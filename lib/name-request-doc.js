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
        children: [new TextRun({ text: line, bold: opts.bold || false })],
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
    // Row 0: Хүсэлт гаргагч
    new TableRow({
      children: [
        makeCell('1.', 1),
        makeCell(
          'Хүсэлт /өргөдөл/ гаргагчийн мэдээлэл /Иргэн, Аж ахуйн нэгж, Төрийн байгууллага, Төрийн бус байгууллага болон бусад/',
          5
        ),
        makeCell(
          'Овог, нэр: Цэрэндорж Орлого\nРД: ОА92103007\nОршин суугаа хаяг: \nУтас: 99913033\nФакс: \nИ-мэйл: orlogo.ts@gazar.gov.mn\nГарын үсэг:\nХавсралт баримтын хуудасны тоо: ____________\nОгноо: _______________',
          9
        ),
      ],
    }),
    // Row 1: Санал болгож буй нэр — 1st (rowSpan for id and label)
    new TableRow({
      children: [
        makeCell('2.', 1, 2),
        makeCell('Санал болгож буй газар зүйн нэр', 3, 2),
        makeCell('1 дэх нэр', 2),
        makeCell(s(row[3]), 9),
      ],
    }),
    // Row 2: 2nd name (cols 0-3 covered by rowSpan)
    new TableRow({
      children: [makeCell('2 дахь нэр', 2), makeCell('', 9)],
    }),
    // Row 3: Нэрний гарал үүсэл
    new TableRow({
      children: [
        makeCell('3.', 1),
        makeCell('Нэрний гарал үүсэл', 5),
        makeCell('Ο Шинээр бий болсон газар зүйн объект\nΟ Газар зүйн уламжлалт нэр', 9),
      ],
    }),
    // Row 4: Дэвсгэр нэр
    new TableRow({
      children: [
        makeCell('4.', 1),
        makeCell('Дэвсгэр нэр /ам, булаг, гол, нуур, уул... гэх мэт/', 5),
        makeCell(s(row[5]), 9),
      ],
    }),
    // Row 5: Аймаг, сум, баг
    new TableRow({
      children: [
        makeCell('5.', 1),
        makeCell('Аймаг, нийслэл, сум, дүүрэг, баг, хорооны нэр, дугаар', 5),
        makeCell(s(row[18]), 9),
      ],
    }),
    // Row 6: Алслагдах зай
    new TableRow({
      children: [
        makeCell('6.', 1),
        makeCell(
          'Хамгийн ойр орших хот, суурин газраас алслагдах зай, километрээр /аль зүгт байрлахыг тодорхой бичих/.',
          5
        ),
        makeCell(s(row[19]), 9),
      ],
    }),
    // Row 7: Солбицол
    new TableRow({
      children: [
        makeCell('7.', 1),
        makeCell('Газар зүйн нэрийн солбицол /градус, минут, секунд/', 5),
        makeCell(coord, 9),
      ],
    }),
    // Row 8: Хэрэглэгдэж буй хугацаа
    new TableRow({
      children: [
        makeCell('8.', 1),
        makeCell(
          'Шинээр бий болсон объектод өгөх нэр, уламжлалт газар зүйн нэрийн хэрэглэгдэж буй хугацаа /жилээр/',
          5
        ),
        makeCell(
          'Ο 50-иас дээш жил /хуучин нэр/\nΟ 10-50 хүртэлх жил /харьцангуй хуучин нэр/\nΟ 10 хүртэлх жил /шинэ нэр/',
          9
        ),
      ],
    }),
    // Row 9: Мэдээллээр хангагч иргэн
    new TableRow({
      children: [
        makeCell('9.', 1),
        makeCell('Нэрийн талаар мэдээллээр хангагч иргэн, хуулийн этгээдийн мэдээлэл', 5),
        makeCell(
          'Овог, нэр: Б.Сэргэлэн\nРегистрийн дугаар: БД60051471\nХаяг: Дундговь аймаг, Луус сум 1-р баг Наран\nУтас:  88261547\nИ-мэйл: sergelenb@gmail.com',
          9
        ),
      ],
    }),
    // Row 10: Зөвлөмж
    new TableRow({
      children: [
        makeCell('10.', 1),
        makeCell('Эрх бүхий байгууллага болон орон нутгийн зөвлөлийн зөвлөмж', 5),
        makeCell(
          '1.Сумын ГЗНСЗ-ийн хурлын шийдвэр\n2.Сумын ИТХ-ын тогтоол\n3.Аймгийн ГЗНСЗ-ийн хурлын шийдвэр\n4.Аймгийн ИТХ-ын тогтоол\n5.ГЗБГЗЗГ-ын ГЗНЗ-ийн хурлын шийдвэр\n6.Газар зүйн нэрийн Үндэсний зөвлөлийн зөвлөмж\n7.Засгийн газар\n8.Үндэсний аюулгүй байдлын зөвлөл\n9.Улсын Их Хурлын тогтоол',
          9
        ),
      ],
    }),
    // Row 11: Гэрэл зураг header (rowSpan for col 0)
    new TableRow({
      children: [makeCell('11.', 1, 2), makeCell('Гэрэл зураг', 14)],
    }),
    // Row 12: Гэрэл зураг description
    new TableRow({
      children: [
        makeCell(
          'Тайлбар: 1-8 ширхэг гэрэл зураг оруулах /зураг дарсан зүг, чиг бичих/;\n/Жишээ нь: Зүүн зүгээс эсвэл Зүүн урд зүгээс гэх мэтээр бичнэ/.',
          14
        ),
      ],
    }),
    // Row 13: Байршлын зураг header (rowSpan for col 0)
    new TableRow({
      children: [makeCell('12.', 1, 2), makeCell('Байршлын зураг', 14)],
    }),
    // Row 14: Байршлын зураг description
    new TableRow({
      children: [
        makeCell(
          '/Газар зүйн нэрийн зураг болон сумын бүдүүвч зураг дээр харагдах байдал/',
          14
        ),
      ],
    }),
  ];
}

export async function generateNameRequest(excelRows) {
  const sections = excelRows.map((row, i) => ({
    properties: {
      type: i === 0 ? SectionType.CONTINUOUS : SectionType.NEXT_PAGE,
    },
    children: [
      new Paragraph({ children: [new TextRun('Хавсралт 1')], alignment: AlignmentType.RIGHT }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Газар зүйн нэрийг шинээр өгөх хүсэлтийн маягт /өргөдөл/',
            bold: true,
            size: 26,
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 100 },
      }),
      new Paragraph({
        children: [new TextRun('Зөвхөн албан хэрэгцээнд:')],
        alignment: AlignmentType.RIGHT,
      }),
      new Table({
        rows: buildRows(row),
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: Array(15).fill(600),
      }),
    ],
  }));

  const doc = new Document({ sections });
  return Packer.toBuffer(doc);
}
