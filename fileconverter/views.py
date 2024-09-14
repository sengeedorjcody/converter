from django.shortcuts import render
from django.http import HttpResponse
from docx import Document
from .forms import UploadFileForm
from .forms import ChangeRequestFileForm
import openpyxl
from docx.shared import Cm, Inches
#
from django.shortcuts import render
from django.http import HttpResponse
from openpyxl import load_workbook
from docx import Document
from docx.shared import RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.text import WD_ALIGN_PARAGRAPH
from django.views.decorators.csrf import csrf_exempt

PAGE_WIDTH_INCHES = 8.5

# def set_column_widths(table):
#     # Calculate column widths in inches
#     col_widths_inches = [PAGE_WIDTH_INCHES * 0.4] + [PAGE_WIDTH_INCHES * 0.2] * 3
#
#     # Convert inches to Cm
#     col_widths_cm = [Inches(width).cm for width in col_widths_inches]
#
#     # Set column widths
#     for idx, width in enumerate(col_widths_cm):
#         table.columns[idx].width = Cm(width)

def handle_uploaded_file(f):
    wb = openpyxl.load_workbook(f)
    ws = wb.active
    doc = Document()
    static_values = ['Газар зүйн нэр /монгол, латин галиг/',
                     'Газар зүйн нэрийн дахин давтагдашгүй дугаар',
                     'Газар зүйн нэрийн гарал, үүсэл',
                     'Газар зүйн нэрийн төрөл /дэвсгэр нэр/',
                     'Харьяалагдах аймаг, сум, баг',
                     'Газар зүйн нэрийн ерөнхий байрлал, тайлбар',
                     'Газар зүйн нэрийн солбицол, UTM, 48-р бүс',
                     'Газар зүйн нэрийн орших 1:25 000-ны масштабтай байр зүйн зургийн нэрлэвэр',
                     '',
                     'Бусад газрын зурагт зөв, ижил хэрэглэгдсэн байдал',
                     'Газар зүйн нэрийг баталгаажуулсан актын нэр, дугаар, огноо',
                     'Өөрчлөлт орсон эсэх, шалтгаан',
                     'Газар зүйн нэрийн байршлын зураг',
                     '',
                     'Газар зүйн нэрийн тодруулалтын үеийн нотлох баримт:',
                     'Газар зүйн нэрийг тодруулсан иргэн, хуулийн этгээд',
                     'Газар зүйн нэрийг тодруулсан газарчин /орон нутгийн/',
                     'Газар зүйн нэрийн хувийн хэрэг хөтөлсөн:',
                     ]
    static_values_second = [
'',
'',
'Уламжлалт нэр',
'худаг',
'',
'Сумын төвөөс хойшоо 27.51 км',
'',
'',
'',
'',
'Сумын иргэдийн төлөөлөгчдийн ...-хурлаар дэмжигдсэн .',
'',
'',
'',
"""Аудио, видео бичлэг: □ 
Тэмдэглэл:    □        Фото зураг:   □""",
"""“Инженер геодези” ХХК-ны инженер:
МУ-ын зөвлөх инженер Д.Оюунчимэг
Инженер: Э.Ануун, Н.Бумчин""",
"""Н.Очирваань, багийн өндөр настан
Э.Эрдэнэтунгалаг, газрын даамал """,
'/2024 оны 05-р сарын 15-ны өдөр/',
]
    for row in ws.iter_rows(min_row=3, values_only=True):
        # doc.add_heading('', 0)
        title = doc.add_paragraph('')
        run = title.add_run('Газар зүйн нэрийн хувийн хэрэг')
        run.bold = True
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        table = doc.add_table(rows=len(row), cols=10)
        # set_column_widths(table)
        table.style = 'Table Grid'
        table.autofit = False
        total_width_cm = 20  # Example total width for table
        table.columns[0].width = Cm(total_width_cm * 0.4)
        # table.columns[0].width = Cm(10)  # Adjust width as needed
        # table.columns[1].width = Cm(2)  # Adjust width as needed
        # table.columns[2].width = Cm(2)  # Adjust width as needed
        # table.columns[3].width = Cm(2)  # Adjust width as needed
        for i, static_value in enumerate(static_values):
            row_cells = table.rows[i].cells
            if i != 12:
                row_cells[0].text = str(static_value) if static_value is not None else ''
                horizantal = table.cell(i, 0).merge(table.cell(i, 3))
                if i == 8:
                    horizantal.merge(table.cell(i+1, 3))
            if i == 0:
                row_cells[4].text = (str(row[2]) if row[2] is not None else '')
                row_cells[7].text = str(row[11]) if row[11] is not None else ''
                table.cell(i, 4).merge(table.cell(i, 6))
                table.cell(i, 7).merge(table.cell(i, 9))
            elif i == 1:
                row_cells[4].text = (str(row[10]) if row[10] is not None else '')
                table.cell(i, 4).merge(table.cell(i, 9))
            elif i == 4:
                row_cells[4].text = (str(row[9]) if row[9] is not None else '')
                table.cell(i, 4).merge(table.cell(i, 9))
            elif i == 6:
                row_cells[4].text = "1.Өргөрөг: " + (str(row[12]) if row[12] is not None else '') + "\n" + "Уртраг: " + (str(row[13]) if row[13] is not None else '')
                table.cell(i, 4).merge(table.cell(i, 9))
            elif i == 7:
                row_cells[4].text = str(row[4]) if row[4] is not None else ''
                table.cell(i, 4).merge(table.cell(i, 5))
                row_cells[6].text = "Нэрийн зургийн индекс"
                table.cell(i, 6).merge(table.cell(i, 7))
                row_cells[8].text = str(row[3]) if row[3] is not None else ''
                table.cell(i, 8).merge(table.cell(i, 9))
            elif i == 8:
                row_cells[4].text = '1:25000 зурагт'
                row_cells[7].text = "1:100 000 зурагт"
                table.cell(i, 4).merge(table.cell(i, 6))
                table.cell(i, 7).merge(table.cell(i, 9))
            elif i == 9:
                row_cells[4].text = 'Үзүүлсэн'
                row_cells[7].text = "Үзүүлсэн"
                table.cell(i, 4).merge(table.cell(i, 6))
                table.cell(i, 7).merge(table.cell(i, 9))
            elif i == 10:
                paragraph = row_cells[4].paragraphs[0]
                run = paragraph.add_run('Сумын иргэдийн төлөөлөгчдийн ...-хурлаар дэмжигдсэн.')
                run.font.color.rgb = RGBColor(255, 0, 0)
                table.cell(i, 4).merge(table.cell(i, 9))
            elif i == 12:
                # row_cells[0].text = str(static_value) if static_value is not None else ''
                paragraph = row_cells[0].paragraphs[0]
                run = paragraph.add_run(str(static_value))
                run.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                table.cell(i, 0).merge(table.cell(i, 9))
            else:
                row_cells[4].text = (str(static_values_second[i]) if static_values_second[i] is not None else '')
                table.cell(i, 4).merge(table.cell(i, 9))

            # row_cells = table.rows[i].cells
            # row_cells[0].text = f'Static Title {i + 1}'  # Adjust static titles as needed
            # row_cells[1].text = str(cell_value) if cell_value is not None else ''
        doc.add_page_break()


    # Save the Word document to a temporary location
    doc_path = 'output.docx'
    doc.save(doc_path)
    return doc_path

def handle_change_request_file(f):
    wb = openpyxl.load_workbook(f)
    ws = wb.active
    doc = Document()
    static_values = [
        {
            "id": "1.",
            "first": "Хүсэлт /өргөдөл/ гаргагчийн мэдээлэл /Иргэн, Аж ахуйн нэгж, Төрийн байгууллага, Төрийн бус байгууллага болон бусад/",
            "second": "Овог, нэр: Цэрэндорж Орлого\nРД: ОА92103007\nОршин суугаа хаяг: \nУтас: 99913033\nФакс: \nИ-мэйл: orlogo.ts@gazar.gov.mn\nГарын үсэг:\nХавсралт баримтын хуудасны тоо: ____________\nОгноо: _______________"
        },
        {
            "id": "2.",
            "first": "Санал болгож буй газар зүйн нэр",
            "second": "",
            "third": "1 дэх нэр",
            "fourth": "2 дахь нэр"
        },
        {
            "id": "2.",
            "first": "Санал болгож буй газар зүйн нэр",
            "second": "",
            "third": "1 дэх нэр",
            "fourth": "2 дахь нэр"
        },
        {
            "id": "3.",
            "first": "Нэрний гарал үүсэл",
            "second": "Ο Шинээр бий болсон газар зүйн объект\nΟ Газар зүйн уламжлалт нэр",
        },
        {
            "id": "4.",
            "first": "Дэвсгэр нэр /ам, булаг, гол, нуур, уул... гэх мэт/",
            "second": "",
        },
        {
            "id": "5.",
            "first": "Аймаг, нийслэл, сум, дүүрэг, баг, хорооны нэр, дугаар",
            "second": "",
        },
        {
            "id": "6.",
            "first": "Хамгийн ойр орших хот, суурин газраас алслагдах зай, километрээр /аль зүгт байрлахыг тодорхой бичих/.",
            "second": "",
        },
        {
            "id": "7.",
            "first": "Газар зүйн нэрийн солбицол /градус, минут, секунд/",
            "second": "",
        },
        {
            "id": "8.",
            "first": "Шинээр бий болсон объектод өгөх нэр, уламжлалт газар зүйн нэрийн хэрэглэгдэж буй хугацаа /жилээр/",
            "second": "Ο 50-иас дээш жил /хуучин нэр/\nΟ 10-50 хүртэлх жил /харьцангуй хуучин нэр/\nΟ 10 хүртэлх жил /шинэ нэр/",
        },
        {
            "id": "9.",
            "first": "Нэрийн талаар мэдээллээр хангагч иргэн, хуулийн этгээдийн мэдээлэл",
            "second": "Овог, нэр: Б.Сэргэлэн\nРегистрийн дугаар: БД60051471\nХаяг: Дундговь аймаг, Луус сум 1-р баг Наран\nУтас:  88261547\nИ-мэйл: sergelenb@gmail.com",
        },
        {
            "id": "10.",
            "first": "Эрх бүхий байгууллага болон орон нутгийн зөвлөлийн зөвлөмж",
            "second": "1.Сумын ГЗНСЗ-ийн хурлын шийдвэр\n2.Сумын ИТХ-ын тогтоол\n3.Аймгийн ГЗНСЗ-ийн хурлын шийдвэр\n4.Аймгийн ИТХ-ын тогтоол\n5.ГЗБГЗЗГ-ын ГЗНЗ-ийн хурлын шийдвэр\n6.Газар зүйн нэрийн Үндэсний зөвлөлийн зөвлөмж\n7.Засгийн газар\n8.Үндэсний аюулгүй байдлын зөвлөл\n9.Улсын Их Хурлын тогтоол",
        },
        {
            "id": "11.",
            "first": "Гэрэл зураг",
            "second": "",
            "third": "",
            "description": "Тайлбар: 1-8 ширхэг гэрэл зураг оруулах /зураг дарсан зүг, чиг бичих/;\n/Жишээ нь: Зүүн зүгээс эсвэл Зүүн урд зүгээс гэх мэтээр бичнэ/.",
        },
        {
            "id": "11.",
            "first": "Гэрэл зураг",
            "second": "",
            "third": "",
            "description": "Тайлбар: 1-8 ширхэг гэрэл зураг оруулах /зураг дарсан зүг, чиг бичих/;\n/Жишээ нь: Зүүн зүгээс эсвэл Зүүн урд зүгээс гэх мэтээр бичнэ/.",
        },
        {
            "id": "12.",
            "first": "Байршлын зураг",
            "second": "",
            "description": "/Газар зүйн нэрийн зураг болон сумын бүдүүвч зураг дээр харагдах байдал/"
        },
        {
            "id": "12.",
            "first": "Байршлын зураг",
            "second": "",
            "description": "/Газар зүйн нэрийн зураг болон сумын бүдүүвч зураг дээр харагдах байдал/"
        },
    ]

    change_request_values = [
        {
            "id": "1.",
            "first": "Хүсэлт /өргөдөл/ гаргагчийн мэдээлэл /Иргэн, Аж ахуйн нэгж, Төрийн байгууллага, Төрийн бус байгууллага болон бусад/",
            "second": "Овог, нэр: \nРД: \nОршин суугаа хаяг: \nУтас: \nФакс: \n\nИ-мэйл: \nГарын үсэг:\nХавсралт баримтын хуудасны тоо: _______\nОгноо: _______________"
        },
        {
            "id": "2.",
            "first": "Батлагдсан газар зүйн нэр",
            "second": "",
            "third": "оноосон нэр",
            "fourth": "дэвсгэр нэр"
        },
        {
            "id": "2.",
            "first": "Батлагдсан газар зүйн нэр",
            "second": "",
            "third": "оноосон нэр",
            "fourth": "дэвсгэр нэр"
        },
        {
            "id": "3.",
            "first": "Батлагдсан огноо:",
            "second": "2003 оны 10 дугаар сарын 31-ний өдөр",
        },
        {
            "id": "4.",
            "first": "Батлагдсан тогтоол, шийдвэрийн дугаар",
            "second": "42 дугаар тогтоол",
        },
        {
            "id": "5.",
            "first": "Баталсан этгээдийн нэр",
            "second": "Улсын Их Хурал",
        },
        {
            "id": "6.",
            "first": "Батлагдсан газар зүйн нэрийн байршил",
            "second": "○ ЗЗНДН-ийн доторх байрлал\n○ Хилийн зааг",
        },
        {
            "id": "7.",
            "first": "Газар зүйн нэрийг өөрчилж буй шалтгаан /тайлбар бичих/",
            "second": "○ Монгол хэлнээс өөр хэл дээр батлагдсан\n○ Үг, үсгийн алдаатай батлагдсан\n○ Адил төрлийн хэд хэдэн объектын ижил нэр нь зам, тээвэр, харилцаа холбоо, бусад байгууллагын ажилд хүндрэл учруулахаар байвал\n○ уугуул иргэд нь нэрлэж заншсан уламжлалт нэрийг сэргээх хүсэлт тавьсан\n○ тухайн объектын мөн чанарт тохирохгүй, этгээд хэллэгээр нэрлэгдсэн байвал",
        },
        {
            "id": "8.",
            "first": "Санал болгож буй нэр",
            "second": "",
            "third": "1 дэх нэр",
            "fourth": "2 дахь нэр"
        },
        {
            "id": "8.",
            "first": "Санал болгож буй нэр",
            "second": "",
            "third": "1 дэх нэр",
            "fourth": "2 дахь нэр"
        },
        {
            "id": "9.",
            "first": "Нэрний гарал үүсэл, утга, хэл, ямар нэрнээс үүсэлтэй талаарх тэмдэглэл",
            "second": "",
        },
        {
            "id": "10.",
            "first": "Аймаг, нийслэл, сум, дүүрэг, баг, хорооны нэр, дугаар.",
            "second": "1.Сумын ГЗНСЗ-ийн хурлын шийдвэр\n2.Сумын ИТХ-ын тогтоол\n3.Аймгийн ГЗНСЗ-ийн хурлын шийдвэр\n4.Аймгийн ИТХ-ын тогтоол\n5.ГЗБГЗЗГ-ын ГЗНЗ-ийн хурлын шийдвэр\n6.Газар зүйн нэрийн Үндэсний зөвлөлийн зөвлөмж\n7.Засгийн газар\n8.Үндэсний аюулгүй байдлын зөвлөл\n9.Улсын Их Хурлын тогтоол",
        },
        {
            "id": "11.",
            "first": "Хамгийн ойр орших хот, суурин газраас алслагдах зай /ямар чиглэлд, ямар зайнд/. 1:100000-ны масштабтай байр зүйн зургийн нэршил",
            "second": ""
        },
        {
            "id": "12.",
            "first": "Газар зүйн нэрийн солбицол /градус, минут, секунд/.",
            "second": "",
            "description": "/Газар зүйн нэрийн зураг болон сумын бүдүүвч зураг дээр харагдах байдал/"
        },
        {
            "id": "13.",
            "first": "Нэрийн талаар мэдээллээр хангагч иргэн, хуулийн этгээдийн мэдээлэл",
            "second": "Овог, нэр:\nРегистрийн дугаар: \nХаяг: \nУтас: \nИ-мэйл:",
        },
        {
            "id": "14.",
            "first": "Эрх бүхий байгууллага болон орон нутгийн зөвлөлийн зөвлөмж",
            "second": "1.Аймаг, нийслэл, сум, дүүргийн ЗДТГ",
        },
        {
            "id": "15.",
            "first": "Гэрэл зураг /зураг дарсан зүг, чиг/",
            "description": "Тайлбар: 1-8 ширхэг гэрэл зураг \nоруулах /зураг дарсан зүг, чиг \nбичих/;\n/Жишээ нь: \nЗүүн зүгээс эсвэл Зүүн урд зүгээс гэх мэтээр \nбичнэ/.",
        },
        {
            "id": "15.",
            "first": "Гэрэл зураг /зураг дарсан зүг, чиг/",
            "description": "Тайлбар: 1-8 ширхэг гэрэл зураг \nоруулах /зураг дарсан зүг, чиг \nбичих/;\n/Жишээ нь: \nЗүүн зүгээс эсвэл Зүүн урд зүгээс гэх мэтээр \nбичнэ/.",
        },
        {
            "id": "16.",
            "first": "Байршлын зураг",
            "description": "/Газар зүйн нэрийн зураг болон сумын схем зураг дээрх харагдах байдал/",
        },
        {
            "id": "16.",
            "first": "Байршлын зураг",
            "description": "/Газар зүйн нэрийн зураг болон сумын схем зураг дээрх харагдах байдал/",
        },
    ]
    for row in ws.iter_rows(min_row=3, values_only=True):
        # doc.add_heading('', 0)
        # top_right_paragraph = doc.add_paragraph('Хавсралт 1')
        # top_right_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # run = top_right_paragraph.runs[0]

        # title = doc.add_paragraph('')
        # run = title.add_run('Газар зүйн нэрийг шинээр өгөх хүсэлтийн маягт /өргөдөл/')
        # run.bold = True
        # title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # top_right_paragraph = doc.add_paragraph('Зөвхөн албан хэрэгцээнд:')
        # top_right_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # run = top_right_paragraph.runs[0]
        

        # table = doc.add_table(rows=15, cols=15)
        # # set_column_widths(table)
        # table.style = 'Table Grid'
        # table.autofit = False
        # total_width_cm = 20  # Example total width for table
        # table.columns[0].width = Cm(total_width_cm * 0.4)
        # # table.columns[0].width = Cm(10)  # Adjust width as needed
        # # table.columns[1].width = Cm(2)  # Adjust width as needed
        # # table.columns[2].width = Cm(2)  # Adjust width as needed
        # # table.columns[3].width = Cm(2)  # Adjust width as needed
        # for i, static_value in enumerate(static_values):
        #     row_cells = table.rows[i].cells
        #     row_cells[0].text = str(static_value.get("id", ""))
        #     row_cells[1].text = str(static_value.get("first", ""))
        #     if i == 1:
        #         horizantal = table.cell(i, 1).merge(table.cell(i, 3))
        #         horizantal.merge(table.cell(i+1, 3))
        #         table.cell(i, 0).merge(table.cell(i+1, 0))
        #         row_cells[4].text = str(static_value.get("third", ""))
        #         table.cell(i, 4).merge(table.cell(i, 5))
        #         table.cell(i + 1 , 4).merge(table.cell(i+1, 5))
        #     elif i == 2:
        #         row_cells[4].text = str(static_value.get("fourth", ""))
        #     elif i == 11 or i == 13:
        #         table.cell(i, 0).merge(table.cell(i+1, 0))
        #         table.cell(i, 1).merge(table.cell(i, 14))
        #         row_cells[1].text = str(static_value.get("first", ""))
        #     elif i == 12 or i == 14:
        #         table.cell(i, 1).merge(table.cell(i, 14))
        #         row_cells[1].text = str(static_value.get("description", ""))
        #     else:
        #         table.cell(i, 1).merge(table.cell(i, 5)) 
            
        #     if i == 7:
        #         row_cells[6].text = "Өргөрөг: " + (str(row[15]) if str(row[15]) is not None else "") + "\n" + "Уртраг: " + (str(row[16]) if str(row[16]) is not None else "")
        #     else:
        #         row_cells[6].text = str(static_value.get("second", ""))
        #     table.cell(i, 6).merge(table.cell(i, 14))

        # doc.add_page_break()

        top_right_paragraph = doc.add_paragraph('Хавсралт 2')
        top_right_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = top_right_paragraph.runs[0]

        title = doc.add_paragraph('')
        run = title.add_run('Газар зүйн нэрийг өөрчлөх хүсэлтийн маягт /өргөдөл/')
        run.bold = True
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        top_right_paragraph = doc.add_paragraph('Зөвхөн албан хэрэгцээнд:')
        top_right_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = top_right_paragraph.runs[0]
        

        table = doc.add_table(rows=20, cols=15)
        # set_column_widths(table)
        table.style = 'Table Grid'
        table.autofit = False
        total_width_cm = 20  # Example total width for table
        table.columns[0].width = Cm(total_width_cm * 0.4)
        # table.columns[0].width = Cm(10)  # Adjust width as needed
        # table.columns[1].width = Cm(2)  # Adjust width as needed
        # table.columns[2].width = Cm(2)  # Adjust width as needed
        # table.columns[3].width = Cm(2)  # Adjust width as needed
        for i, static_value in enumerate(change_request_values):
            row_cells = table.rows[i].cells
            row_cells[0].text = str(static_value.get("id", ""))
            row_cells[1].text = str(static_value.get("first", ""))
            if i == 1 or i == 8:
                horizantal = table.cell(i, 1).merge(table.cell(i, 3))
                horizantal.merge(table.cell(i+1, 3))
                table.cell(i, 0).merge(table.cell(i+1, 0))
                row_cells[4].text = str(static_value.get("third", ""))
                table.cell(i, 4).merge(table.cell(i, 5))
                table.cell(i + 1 , 4).merge(table.cell(i+1, 5))
            elif i == 2 or i == 9:
                row_cells[4].text = str(static_value.get("fourth", ""))
            elif i == 16 or i == 18:
                table.cell(i, 0).merge(table.cell(i+1, 0))
                table.cell(i, 1).merge(table.cell(i, 14))
                row_cells[1].text = str(static_value.get("first", ""))
            elif i == 17 or i == 19:
                table.cell(i, 1).merge(table.cell(i, 14))
                row_cells[1].text = str(static_value.get("description", ""))
            else:
                table.cell(i, 1).merge(table.cell(i, 5)) 
            
            row_cells[6].text = str(static_value.get("second", ""))
            table.cell(i, 6).merge(table.cell(i, 14))

        doc.add_page_break()


    # Save the Word document to a temporary location
    doc_path = 'output.docx'
    doc.save(doc_path)
    return doc_path

def handle_name_request_file(f):
    wb = openpyxl.load_workbook(f)
    ws = wb.active
    doc = Document()
    static_values = [
        {
            "id": "1.",
            "first": "Хүсэлт /өргөдөл/ гаргагчийн мэдээлэл /Иргэн, Аж ахуйн нэгж, Төрийн байгууллага, Төрийн бус байгууллага болон бусад/",
            "second": "Овог, нэр: Цэрэндорж Орлого\nРД: ОА92103007\nОршин суугаа хаяг: \nУтас: 99913033\nФакс: \nИ-мэйл: orlogo.ts@gazar.gov.mn\nГарын үсэг:\nХавсралт баримтын хуудасны тоо: ____________\nОгноо: _______________"
        },
        {
            "id": "2.",
            "first": "Санал болгож буй газар зүйн нэр",
            "second": "",
            "third": "1 дэх нэр",
            "fourth": "2 дахь нэр"
        },
        {
            "id": "2.",
            "first": "Санал болгож буй газар зүйн нэр",
            "second": "",
            "third": "1 дэх нэр",
            "fourth": "2 дахь нэр"
        },
        {
            "id": "3.",
            "first": "Нэрний гарал үүсэл",
            "second": "Ο Шинээр бий болсон газар зүйн объект\nΟ Газар зүйн уламжлалт нэр",
        },
        {
            "id": "4.",
            "first": "Дэвсгэр нэр /ам, булаг, гол, нуур, уул... гэх мэт/",
            "second": "",
        },
        {
            "id": "5.",
            "first": "Аймаг, нийслэл, сум, дүүрэг, баг, хорооны нэр, дугаар",
            "second": "",
        },
        {
            "id": "6.",
            "first": "Хамгийн ойр орших хот, суурин газраас алслагдах зай, километрээр /аль зүгт байрлахыг тодорхой бичих/.",
            "second": "",
        },
        {
            "id": "7.",
            "first": "Газар зүйн нэрийн солбицол /градус, минут, секунд/",
            "second": "",
        },
        {
            "id": "8.",
            "first": "Шинээр бий болсон объектод өгөх нэр, уламжлалт газар зүйн нэрийн хэрэглэгдэж буй хугацаа /жилээр/",
            "second": "Ο 50-иас дээш жил /хуучин нэр/\nΟ 10-50 хүртэлх жил /харьцангуй хуучин нэр/\nΟ 10 хүртэлх жил /шинэ нэр/",
        },
        {
            "id": "9.",
            "first": "Нэрийн талаар мэдээллээр хангагч иргэн, хуулийн этгээдийн мэдээлэл",
            "second": "Овог, нэр: Б.Сэргэлэн\nРегистрийн дугаар: БД60051471\nХаяг: Дундговь аймаг, Луус сум 1-р баг Наран\nУтас:  88261547\nИ-мэйл: sergelenb@gmail.com",
        },
        {
            "id": "10.",
            "first": "Эрх бүхий байгууллага болон орон нутгийн зөвлөлийн зөвлөмж",
            "second": "1.Сумын ГЗНСЗ-ийн хурлын шийдвэр\n2.Сумын ИТХ-ын тогтоол\n3.Аймгийн ГЗНСЗ-ийн хурлын шийдвэр\n4.Аймгийн ИТХ-ын тогтоол\n5.ГЗБГЗЗГ-ын ГЗНЗ-ийн хурлын шийдвэр\n6.Газар зүйн нэрийн Үндэсний зөвлөлийн зөвлөмж\n7.Засгийн газар\n8.Үндэсний аюулгүй байдлын зөвлөл\n9.Улсын Их Хурлын тогтоол",
        },
        {
            "id": "11.",
            "first": "Гэрэл зураг",
            "second": "",
            "third": "",
            "description": "Тайлбар: 1-8 ширхэг гэрэл зураг оруулах /зураг дарсан зүг, чиг бичих/;\n/Жишээ нь: Зүүн зүгээс эсвэл Зүүн урд зүгээс гэх мэтээр бичнэ/.",
        },
        {
            "id": "11.",
            "first": "Гэрэл зураг",
            "second": "",
            "third": "",
            "description": "Тайлбар: 1-8 ширхэг гэрэл зураг \nоруулах /зураг дарсан зүг, чиг\nбичих/;\n/Жишээ нь: Зүүн зүгээс эсвэл \nЗүүн урд зүгээс гэх мэтээр \nбичнэ/.",
        },
        {
            "id": "12.",
            "first": "Байршлын зураг",
            "second": "",
            "description": "/Газар зүйн нэрийн зураг болон сумын бүдүүвч зураг дээр харагдах байдал/"
        },
        {
            "id": "12.",
            "first": "Байршлын зураг",
            "second": "",
            "description": "\n\n\n/Газар зүйн нэрийн зураг болон сумын бүдүүвч зураг дээр харагдах байдал/"
        },
    ]

    # change_request_values = [
    #     {
    #         "id": "1.",
    #         "first": "Хүсэлт /өргөдөл/ гаргагчийн мэдээлэл /Иргэн, Аж ахуйн нэгж, Төрийн байгууллага, Төрийн бус байгууллага болон бусад/",
    #         "second": "Овог, нэр: \nРД: \nОршин суугаа хаяг: \nУтас: \nФакс: \n\nИ-мэйл: \nГарын үсэг:\nХавсралт баримтын хуудасны тоо: _______\nОгноо: _______________"
    #     },
    #     {
    #         "id": "2.",
    #         "first": "Батлагдсан газар зүйн нэр",
    #         "second": "",
    #         "third": "оноосон нэр",
    #         "fourth": "дэвсгэр нэр"
    #     },
    #     {
    #         "id": "2.",
    #         "first": "Батлагдсан газар зүйн нэр",
    #         "second": "",
    #         "third": "оноосон нэр",
    #         "fourth": "дэвсгэр нэр"
    #     },
    #     {
    #         "id": "3.",
    #         "first": "Батлагдсан огноо:",
    #         "second": "2003 оны 10 дугаар сарын 31-ний өдөр",
    #     },
    #     {
    #         "id": "4.",
    #         "first": "Батлагдсан тогтоол, шийдвэрийн дугаар",
    #         "second": "42 дугаар тогтоол",
    #     },
    #     {
    #         "id": "5.",
    #         "first": "Баталсан этгээдийн нэр",
    #         "second": "Улсын Их Хурал",
    #     },
    #     {
    #         "id": "6.",
    #         "first": "Батлагдсан газар зүйн нэрийн байршил",
    #         "second": "○ ЗЗНДН-ийн доторх байрлал\n○ Хилийн зааг",
    #     },
    #     {
    #         "id": "7.",
    #         "first": "Газар зүйн нэрийг өөрчилж буй шалтгаан /тайлбар бичих/",
    #         "second": "○ Монгол хэлнээс өөр хэл дээр батлагдсан\n○ Үг, үсгийн алдаатай батлагдсан\n○ Адил төрлийн хэд хэдэн объектын ижил нэр нь зам, тээвэр, харилцаа холбоо, бусад байгууллагын ажилд хүндрэл учруулахаар байвал\n○ уугуул иргэд нь нэрлэж заншсан уламжлалт нэрийг сэргээх хүсэлт тавьсан\n○ тухайн объектын мөн чанарт тохирохгүй, этгээд хэллэгээр нэрлэгдсэн байвал",
    #     },
    #     {
    #         "id": "8.",
    #         "first": "Санал болгож буй нэр",
    #         "second": "",
    #         "third": "1 дэх нэр",
    #         "fourth": "2 дахь нэр"
    #     },
    #     {
    #         "id": "8.",
    #         "first": "Санал болгож буй нэр",
    #         "second": "",
    #         "third": "1 дэх нэр",
    #         "fourth": "2 дахь нэр"
    #     },
    #     {
    #         "id": "9.",
    #         "first": "Нэрний гарал үүсэл, утга, хэл, ямар нэрнээс үүсэлтэй талаарх тэмдэглэл",
    #         "second": "",
    #     },
    #     {
    #         "id": "10.",
    #         "first": "Аймаг, нийслэл, сум, дүүрэг, баг, хорооны нэр, дугаар.",
    #         "second": "1.Сумын ГЗНСЗ-ийн хурлын шийдвэр\n2.Сумын ИТХ-ын тогтоол\n3.Аймгийн ГЗНСЗ-ийн хурлын шийдвэр\n4.Аймгийн ИТХ-ын тогтоол\n5.ГЗБГЗЗГ-ын ГЗНЗ-ийн хурлын шийдвэр\n6.Газар зүйн нэрийн Үндэсний зөвлөлийн зөвлөмж\n7.Засгийн газар\n8.Үндэсний аюулгүй байдлын зөвлөл\n9.Улсын Их Хурлын тогтоол",
    #     },
    #     {
    #         "id": "11.",
    #         "first": "Хамгийн ойр орших хот, суурин газраас алслагдах зай /ямар чиглэлд, ямар зайнд/. 1:100000-ны масштабтай байр зүйн зургийн нэршил",
    #         "second": ""
    #     },
    #     {
    #         "id": "12.",
    #         "first": "Газар зүйн нэрийн солбицол /градус, минут, секунд/.",
    #         "second": "",
    #         "description": "/Газар зүйн нэрийн зураг болон сумын бүдүүвч зураг дээр харагдах байдал/"
    #     },
    #     {
    #         "id": "13.",
    #         "first": "Нэрийн талаар мэдээллээр хангагч иргэн, хуулийн этгээдийн мэдээлэл",
    #         "second": "Овог, нэр:\nРегистрийн дугаар: \nХаяг: \nУтас: \nИ-мэйл:",
    #     },
    #     {
    #         "id": "14.",
    #         "first": "Эрх бүхий байгууллага болон орон нутгийн зөвлөлийн зөвлөмж",
    #         "second": "1.Аймаг, нийслэл, сум, дүүргийн ЗДТГ",
    #     },
    #     {
    #         "id": "15.",
    #         "first": "Гэрэл зураг /зураг дарсан зүг, чиг/",
    #         "description": "Тайлбар: 1-8 ширхэг гэрэл зураг \nоруулах /зураг дарсан зүг, чиг \nбичих/;\n/Жишээ нь: \nЗүүн зүгээс эсвэл Зүүн урд зүгээс гэх мэтээр \nбичнэ/.",
    #     },
    #     {
    #         "id": "15.",
    #         "first": "Гэрэл зураг /зураг дарсан зүг, чиг/",
    #         "description": "Тайлбар: 1-8 ширхэг гэрэл зураг \nоруулах /зураг дарсан зүг, чиг \nбичих/;\n/Жишээ нь: \nЗүүн зүгээс эсвэл Зүүн урд зүгээс гэх мэтээр \nбичнэ/.",
    #     },
    #     {
    #         "id": "16.",
    #         "first": "Байршлын зураг",
    #         "description": "/Газар зүйн нэрийн зураг болон сумын схем зураг дээрх харагдах байдал/",
    #     },
    #     {
    #         "id": "16.",
    #         "first": "Байршлын зураг",
    #         "description": "/Газар зүйн нэрийн зураг болон сумын схем зураг дээрх харагдах байдал/",
    #     },
    # ]
    for row in ws.iter_rows(min_row=3, values_only=True):
        # doc.add_heading('', 0)
        top_right_paragraph = doc.add_paragraph('Хавсралт 1')
        top_right_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = top_right_paragraph.runs[0]

        title = doc.add_paragraph('')
        run = title.add_run('Газар зүйн нэрийг шинээр өгөх хүсэлтийн маягт /өргөдөл/')
        run.bold = True
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        top_right_paragraph = doc.add_paragraph('Зөвхөн албан хэрэгцээнд:')
        top_right_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = top_right_paragraph.runs[0]

        table = doc.add_table(rows=15, cols=15)
        # set_column_widths(table)
        table.style = 'Table Grid'
        table.autofit = False
        total_width_cm = 20  # Example total width for table
        table.columns[0].width = Cm(total_width_cm * 0.4)
        # table.columns[0].width = Cm(10)  # Adjust width as needed
        # table.columns[1].width = Cm(2)  # Adjust width as needed
        # table.columns[2].width = Cm(2)  # Adjust width as needed
        # table.columns[3].width = Cm(2)  # Adjust width as needed
        for i, static_value in enumerate(static_values):
            row_cells = table.rows[i].cells
            row_cells[0].text = str(static_value.get("id", ""))
            row_cells[1].text = str(static_value.get("first", ""))
            if i == 1:
                horizantal = table.cell(i, 1).merge(table.cell(i, 3))
                horizantal.merge(table.cell(i+1, 3))
                table.cell(i, 0).merge(table.cell(i+1, 0))
                row_cells[4].text = str(static_value.get("third", ""))
                table.cell(i, 4).merge(table.cell(i, 5))
                table.cell(i + 1 , 4).merge(table.cell(i+1, 5))
            elif i == 2:
                row_cells[4].text = str(static_value.get("fourth", ""))
            elif i == 11 or i == 13:
                table.cell(i, 0).merge(table.cell(i+1, 0))
                table.cell(i, 1).merge(table.cell(i, 14))
                row_cells[1].text = str(static_value.get("first", ""))
            elif i == 12 or i == 14:
                table.cell(i, 1).merge(table.cell(i, 14))
                row_cells[1].text = str(static_value.get("description", ""))
            else:
                table.cell(i, 1).merge(table.cell(i, 5)) 
            
            if i == 7:
                row_cells[6].text = "Өргөрөг: " + (str(row[15]) if str(row[15]) is not None else "") + "\n" + "Уртраг: " + (str(row[16]) if str(row[16]) is not None else "")
            elif i == 1:
                row_cells[6].text = (str(row[2]) if str(row[2]) is not None else "")
            else:
                row_cells[6].text = str(static_value.get("second", ""))
            table.cell(i, 6).merge(table.cell(i, 14))

        doc.add_page_break()

        # top_right_paragraph = doc.add_paragraph('Хавсралт 2')
        # top_right_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # run = top_right_paragraph.runs[0]

        # title = doc.add_paragraph('')
        # run = title.add_run('Газар зүйн нэрийг өөрчлөх хүсэлтийн маягт /өргөдөл/')
        # run.bold = True
        # title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # top_right_paragraph = doc.add_paragraph('Зөвхөн албан хэрэгцээнд:')
        # top_right_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # run = top_right_paragraph.runs[0]

        # table = doc.add_table(rows=20, cols=15)
        # # set_column_widths(table)
        # table.style = 'Table Grid'
        # table.autofit = False
        # total_width_cm = 20  # Example total width for table
        # table.columns[0].width = Cm(total_width_cm * 0.4)
        # # table.columns[0].width = Cm(10)  # Adjust width as needed
        # # table.columns[1].width = Cm(2)  # Adjust width as needed
        # # table.columns[2].width = Cm(2)  # Adjust width as needed
        # # table.columns[3].width = Cm(2)  # Adjust width as needed
        # for i, static_value in enumerate(change_request_values):
        #     row_cells = table.rows[i].cells
        #     row_cells[0].text = str(static_value.get("id", ""))
        #     row_cells[1].text = str(static_value.get("first", ""))
        #     if i == 1 or i == 8:
        #         horizantal = table.cell(i, 1).merge(table.cell(i, 3))
        #         horizantal.merge(table.cell(i+1, 3))
        #         table.cell(i, 0).merge(table.cell(i+1, 0))
        #         row_cells[4].text = str(static_value.get("third", ""))
        #         table.cell(i, 4).merge(table.cell(i, 5))
        #         table.cell(i + 1 , 4).merge(table.cell(i+1, 5))
        #     elif i == 2 or i == 9:
        #         row_cells[4].text = str(static_value.get("fourth", ""))
        #     elif i == 16 or i == 18:
        #         table.cell(i, 0).merge(table.cell(i+1, 0))
        #         table.cell(i, 1).merge(table.cell(i, 14))
        #         row_cells[1].text = str(static_value.get("first", ""))
        #     elif i == 17 or i == 19:
        #         table.cell(i, 1).merge(table.cell(i, 14))
        #         row_cells[1].text = str(static_value.get("description", ""))
        #     else:
        #         table.cell(i, 1).merge(table.cell(i, 5)) 
            
        #     row_cells[6].text = str(static_value.get("second", ""))
        #     table.cell(i, 6).merge(table.cell(i, 14))

        # doc.add_page_break()


    # Save the Word document to a temporary location
    doc_path = 'output.docx'
    doc.save(doc_path)
    return doc_path

def home(request):
    return render(request, 'base.html')

@csrf_exempt
def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            doc_path = handle_uploaded_file(request.FILES['file'])
            with open(doc_path, 'rb') as fh:
                response = HttpResponse(fh.read(),
                                        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                response['Content-Disposition'] = 'inline; filename=' + doc_path
                return response
    else:
        form = UploadFileForm()
    return render(request, 'fileconverter/upload.html', {'form': form})

@csrf_exempt
def change_request(request):
    if request.method == 'POST':
        form = ChangeRequestFileForm(request.POST, request.FILES)
        if form.is_valid():
            doc_path = handle_change_request_file(request.FILES['file'])
            with open(doc_path, 'rb') as fh:
                response = HttpResponse(fh.read(),
                                        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                response['Content-Disposition'] = 'inline; filename=' + doc_path
                return response
    else:
        form = ChangeRequestFileForm()
    return render(request, 'fileconverter/change_request.html', {'form': form})

@csrf_exempt
def name_request(request):
    if request.method == 'POST':
        form = ChangeRequestFileForm(request.POST, request.FILES)
        if form.is_valid():
            doc_path = handle_name_request_file(request.FILES['file'])
            with open(doc_path, 'rb') as fh:
                response = HttpResponse(fh.read(),
                                        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                response['Content-Disposition'] = 'inline; filename=' + doc_path
                return response
    else:
        form = ChangeRequestFileForm()
    return render(request, 'fileconverter/name_request.html', {'form': form})

def excel_to_word_view(request):
    if request.method == 'POST' and request.FILES['excel_file']:
        excel_file = request.FILES['excel_file']
        wb = load_workbook(excel_file)
        sheet = wb.active

        document = Document()

        # Iterate through Excel rows and convert each to a table
        for row in sheet.iter_rows(min_row=2):  # assuming the first row is the header
            table = document.add_table(rows=1, cols=len(row))
            hdr_cells = table.rows[0].cells
            for i, cell in enumerate(row):
                hdr_cells[i].text = str(cell.value) if cell.value else ''

                # Apply color if the cell has a fill color
                if cell.fill.start_color.index != '00000000':  # Not default color
                    # hex_color = cell.fill.start_color.rgb[2:]  # Get hex value
                    hex_color = '000'
                    shading_elm_1 = hdr_cells[i]._element
                    shading_elm_1.get_or_add_tcPr().append(
                        parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
                    )

        # Save the document
        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        response['Content-Disposition'] = 'attachment; filename=converted.docx'
        document.save(response)
        return response

    return render(request, 'fileconverter/convert.html')
