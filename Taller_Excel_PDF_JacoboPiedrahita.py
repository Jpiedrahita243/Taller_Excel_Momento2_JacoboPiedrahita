import openpyxl
from tabulate import tabulate
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors


# Abrir el Excel
wb = openpyxl.load_workbook("ventas_tecnologia_v2.xlsx", data_only=True)
sheet = wb.active


# Guardar los productos en una lista
data = []

for row in sheet.iter_rows(min_row=2, values_only=True):
    data.append(list(row))

wb.close()


# Nombres y alineación de las columnas
headers = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (Unid)"]
alignments = ("center", "left", "center", "right", "right")


# Mostrar productos en la consola
def mostrar_tabla(filas):
    if len(filas) == 0:
        print("No se encontraron productos.")
    else:
        print(tabulate(
            filas,
            headers=headers,
            tablefmt="fancy_grid",
            colalign=alignments,
        ))


# Mostrar el catálogo
def ver_catalogo():
    print(f"\n=== CATÁLOGO DE VENTAS ({len(data)} PRODUCTOS) ===")
    mostrar_tabla(data)


# Filtrar productos que cuestan más de $100
def filtrar_precio():
    filtrados = [row for row in data if row[3] > 100]

    print("\n=== PRODUCTOS CON PRECIO MAYOR A $100 ===")
    mostrar_tabla(filtrados)


# Calcular las métricas
def mostrar_resumen():
    total_unidades = sum(row[4] for row in data)
    precio_promedio = sum(row[3] for row in data) / len(data)

    print("\n=== RESUMEN ===")
    print(f"Total de unidades vendidas: {total_unidades}")
    print(f"Precio promedio: ${precio_promedio:,.2f}")


# Buscar un producto por ID
def buscar_id():
    id_buscado = input("Escribe el ID del producto: ")

    try:
        id_buscado = int(id_buscado)
    except ValueError:
        print("El ID debe ser un número.")
        return

    encontrado = [row for row in data if row[0] == id_buscado]

    if len(encontrado) > 0:
        mostrar_tabla(encontrado)
    else:
        print("No se encontró ese ID.")


# Crear el informe PDF
def crear_pdf():
    filtrados = [row for row in data if row[3] > 100]

    total_unidades = sum(row[4] for row in data)
    precio_promedio = sum(row[3] for row in data) / len(data)

    doc = SimpleDocTemplate(
        "reporte_ejecutivo_ventas.pdf",
        pagesize=letter,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=18,
        textColor=colors.HexColor("#1A365D"),
    )

    body_style = ParagraphStyle(
        "BodyStyle",
        parent=styles["Normal"],
        fontSize=9,
    )

    header_style = ParagraphStyle(
        "HeaderStyle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        textColor=colors.whitesmoke,
        alignment=1,
    )

    story = [
        Paragraph("Informe Ejecutivo de Ventas Tecno", title_style),
        Paragraph("Productos con precio mayor a $100", body_style),
        Spacer(1, 12),
    ]

    header_pdf = [
        Paragraph(f"<b>{header}</b>", header_style)
        for header in headers
    ]

    data_pdf = []

    for row in filtrados:
        data_pdf.append([
            Paragraph(str(row[0]), body_style),
            Paragraph(str(row[1]), body_style),
            Paragraph(str(row[2]), body_style),
            Paragraph(f"${row[3]:,.2f}", body_style),
            Paragraph(str(row[4]), body_style),
        ])

    # Agregar el resumen al final de la tabla
    resumen = [
        "",
        Paragraph("<b>RESUMEN</b>", body_style),
        "",
        Paragraph(f"Promedio: ${precio_promedio:,.2f}", body_style),
        Paragraph(f"Unidades: {total_unidades}", body_style),
    ]

    tabla = Table(
        [header_pdf] + data_pdf + [resumen],
        colWidths=[40, 160, 100, 80, 80],
        repeatRows=1,
    )

    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        (
            "ROWBACKGROUNDS",
            (0, 1),
            (-1, -2),
            [colors.white, colors.HexColor("#F7FAFC")],
        ),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D9EAD3")),
    ]))

    story.append(tabla)
    doc.build(story)

    print("Se creó el archivo reporte_ejecutivo_ventas.pdf")


# Mostrar el menú
def mostrar_menu():
    while True:
        print("\n=== MENÚ ===")
        print("1. Ver catálogo completo")
        print("2. Filtrar por precio")
        print("3. Buscar por ID")
        print("4. Crear informe PDF")
        print("5. Salir")

        opcion = input("Elige una opción: ")

        if opcion == "1":
            ver_catalogo()
        elif opcion == "2":
            filtrar_precio()
        elif opcion == "3":
            buscar_id()
        elif opcion == "4":
            crear_pdf()
        elif opcion == "5":
            mostrar_resumen()
            print("Programa terminado.")
            break
        else:
            print("Opción incorrecta. Elige del 1 al 5.")


print(f"Total de registros leídos: {len(data)}")
mostrar_menu()