package com.mycompany.ocxrst.vistas;

import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.AreaBreakType;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import com.itextpdf.layout.properties.VerticalAlignment;
import com.mycompany.ocxrst.config.AppConfig;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * Construye el PDF de una orden de compra sin depender de componentes Swing.
 *
 * <p>La vista solamente transforma sus controles a {@link DatosOrden}; toda
 * la estructura visual, formato y cálculo del documento vive en esta clase.</p>
 */
public final class PDFORDENCOMPRA {

    private static final DeviceRgb AZUL = new DeviceRgb(0, 95, 131);
    private static final DeviceRgb SEPARADOR = new DeviceRgb(42, 210, 201);
    private static final DeviceRgb AZUL_CLARO = new DeviceRgb(204, 229, 239);
    private static final DeviceRgb BLANCO = new DeviceRgb(255, 255, 255);

    private PDFORDENCOMPRA() {
    }

    public static final class Producto {
        public String codigo = "";
        public String descripcion = "";
        public String unidad = "";
        public String cantidad = "";
        public double precioUnitario;
        public double importe;
    }

    public static final class DatosOrden {
        public String documento = "";
        public String numeroOrden = "";
        public String cotizacion = "";
        public Date fecha;
        public String proyecto = "";
        public String idProveedor = "";
        public String proveedor = "";
        public String rfc = "";
        public String contacto = "";
        public String telefono = "";
        public String direccion = "";
        public String correo = "";
        public String pago = "";
        public Date fechaEntrega;
        public Date fechaPago;
        public String tipoCambio = "";
        public String moneda = "MXN";
        public String monedaDescripcion = "";
        public String monedaLeyenda = "";
        public String solicitante = "";
        public String area = "";
        public String usoCfdi = "";
        public String formaPago = "";
        public String metodoPago = "";
        public String elaboro = "";
        public String autorizo = "";
        public String subtotalTexto = "";
        public String ivaTexto = "";
        public String totalTexto = "";
        public double subtotalOriginal;
        public double descuentoImporte;
        public double isrImporte;
        public double descuentoPorcentaje;
        public boolean mostrarDescuento;
        public boolean mostrarIsr;
        public String observaciones = "";
        public String terminos = "";
        public final List<Producto> productos = new ArrayList<>();
        public final Map<String, String> valores = new java.util.HashMap<>();
    }

    /**
     * Genera o reemplaza el archivo indicado y devuelve el mismo archivo.
     *
     * @throws IOException si no puede crear el PDF o leer un recurso requerido
     */
    public static File generar(File destino, DatosOrden datos) throws IOException {
        if (destino == null || datos == null) {
            throw new IllegalArgumentException("El destino y los datos de la orden son obligatorios.");
        }
        File parent = destino.getParentFile();
        if (parent != null) {
            Files.createDirectories(parent.toPath());
        }

        try (PdfWriter writer = new PdfWriter(destino);
                PdfDocument pdf = new PdfDocument(writer)) {
            PdfFont bold = PdfFontFactory.createFont(StandardFonts.HELVETICA_BOLD);
            PdfFont plain = PdfFontFactory.createFont(StandardFonts.HELVETICA);
            Document doc = new Document(pdf, new PageSize(PageSize.LETTER.getWidth(),
                    calcularAltoPrimeraPagina(datos)));
            doc.setMargins(36, 36, 36, 36);

            agregarEncabezado(doc, datos, bold, plain);
            doc.add(lineaSeparadora());
            agregarInformacion(doc, datos, bold, plain);
            agregarProveedor(doc, datos, bold, plain);
            agregarCondiciones(doc, datos, bold, plain);
            agregarSolicitud(doc, datos, bold, plain);
            agregarProductos(doc, datos, bold, plain);
            agregarTotales(doc, datos, bold, plain);
            agregarFactura(doc, datos, bold, plain);
            agregarFirmas(doc, datos, bold, plain);

            pdf.setDefaultPageSize(PageSize.LETTER);
            doc.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
            doc.add(seccion("TÉRMINOS Y CONDICIONES", bold));
            doc.add(new Paragraph(datos.terminos == null || datos.terminos.isBlank()
                    ? "(No se pudo cargar el archivo de términos y condiciones)"
                    : datos.terminos).setFont(plain).setFontSize(8)
                    .setMarginTop(6).setTextAlignment(TextAlignment.JUSTIFIED));
            doc.close();
        }
        return destino;
    }

    private static void agregarEncabezado(Document doc, DatosOrden d, PdfFont bold, PdfFont plain)
            throws IOException {
        Table header = new Table(UnitValue.createPercentArray(new float[]{25, 50, 25}))
                .useAllAvailableWidth().setMarginBottom(3);
        Image logo = new Image(ImageDataFactory.create(
                AppConfig.recursoBytes("IMAGENES/Logo PDF.png"))).scaleToFit(130, 55);
        header.addCell(new Cell().add(logo).setBorder(Border.NO_BORDER).setPadding(4)
                .setVerticalAlignment(VerticalAlignment.MIDDLE));
        header.addCell(new Cell()
                .add(new Paragraph("ORDEN DE COMPRA").setFont(bold).setFontSize(18)
                        .setFontColor(AZUL).setTextAlignment(TextAlignment.CENTER))
                .add(new Paragraph("RED SINERGIA DE TELECOMUNICACIONES, S.A. DE C.V.\n"
                        + "Av. Ocote #10 Arboledas Guadalupe, C.P. 72260. RST140517D48\n"
                        + "TELS.: +52 (01 2222) 278842, (01 55) 7823 1585")
                        .setFont(plain).setFontSize(7).setFontColor(AZUL)
                        .setTextAlignment(TextAlignment.CENTER))
                .setBorder(Border.NO_BORDER).setPadding(4)
                .setVerticalAlignment(VerticalAlignment.MIDDLE));
        header.addCell(new Cell()
                .add(new Paragraph(d.documento).setFont(bold).setFontSize(8)
                        .setTextAlignment(TextAlignment.RIGHT))
                .add(new Paragraph("NO.  " + d.numeroOrden).setFont(bold).setFontSize(9)
                        .setFontColor(AZUL).setTextAlignment(TextAlignment.RIGHT))
                .add(new Paragraph("Cotización: " + d.cotizacion).setFont(plain).setFontSize(9)
                        .setTextAlignment(TextAlignment.RIGHT))
                .setBorder(Border.NO_BORDER).setPadding(4)
                .setVerticalAlignment(VerticalAlignment.MIDDLE));
        doc.add(header);
    }

    private static void agregarInformacion(Document doc, DatosOrden d, PdfFont bold, PdfFont plain) {
        Table table = tablaInformacion();
        fila(table, bold, plain, "FECHA:", fecha(d.fecha), "PROYECTO:", d.proyecto);
        doc.add(table);
    }

    private static void agregarProveedor(Document doc, DatosOrden d, PdfFont bold, PdfFont plain) {
        doc.add(seccion("DATOS DEL PROVEEDOR", bold));
        Table table = tablaInformacion();
        fila(table, bold, plain, "ID PROVEEDOR:", d.idProveedor, "PROVEEDOR:", d.proveedor);
        fila(table, bold, plain, "RFC:", stripHtml(d.rfc), "DIRECCIÓN:", stripHtml(d.direccion));
        fila(table, bold, plain, "CONTACTO:", stripHtml(d.contacto), "CORREO:", d.correo);
        fila(table, bold, plain, "TELÉFONO:", d.telefono, "", "");
        doc.add(table);
    }

    private static void agregarCondiciones(Document doc, DatosOrden d, PdfFont bold, PdfFont plain) {
        doc.add(seccion("CONDICIONES COMERCIALES", bold));
        Table table = tablaInformacion();
        fila(table, bold, plain, "PAGO:", d.pago, "FECHA DE ENTREGA:", fecha(d.fechaEntrega));
        fila(table, bold, plain, "FECHA DE PAGO:", fecha(d.fechaPago), "", "");
        if (d.tipoCambio != null && !d.tipoCambio.isBlank()) {
            fila(table, bold, plain, "TIPO DE CAMBIO:", "1 " + d.moneda + " = "
                    + d.tipoCambio + " MXN", "", "");
        }
        doc.add(table);
    }

    private static void agregarSolicitud(Document doc, DatosOrden d, PdfFont bold, PdfFont plain) {
        doc.add(seccion("DATOS DE LA SOLICITUD", bold));
        Table table = tablaInformacion();
        fila(table, bold, plain, "SOLICITANTE:", d.solicitante, "ÁREA:", d.area);
        doc.add(table);
    }

    private static void agregarProductos(Document doc, DatosOrden d, PdfFont bold, PdfFont plain) {
        doc.add(seccion("PRODUCTOS / SERVICIOS", bold));
        Table table = new Table(UnitValue.createPercentArray(
                new float[]{1.5f, 4f, 2f, 1f, 1.5f, 1.5f}))
                .useAllAvailableWidth().setMarginBottom(2).setKeepTogether(false);
        String[] headers = {"CÓDIGO", "DESCRIPCIÓN", "UNIDAD DE MEDIDA", "CANT.",
            "PRECIO UNIT.", "IMPORTE"};
        TextAlignment[] aligns = {TextAlignment.CENTER, TextAlignment.LEFT, TextAlignment.CENTER,
            TextAlignment.CENTER, TextAlignment.RIGHT, TextAlignment.RIGHT};
        for (int i = 0; i < headers.length; i++) {
            table.addHeaderCell(new Cell().add(new Paragraph(headers[i]).setFont(bold)
                    .setFontSize(8).setFontColor(ColorConstants.WHITE))
                    .setBackgroundColor(AZUL).setTextAlignment(aligns[i]).setPadding(4));
        }
        boolean alternate = false;
        for (Producto p : d.productos) {
            DeviceRgb bg = alternate ? AZUL_CLARO : BLANCO;
            table.addCell(celda(p.codigo, plain, TextAlignment.CENTER, bg));
            table.addCell(celda(p.descripcion, plain, TextAlignment.LEFT, bg));
            table.addCell(celda(p.unidad, plain, TextAlignment.CENTER, bg));
            table.addCell(celda(p.cantidad, plain, TextAlignment.CENTER, bg));
            table.addCell(celdaMoneda(p.precioUnitario, plain, bg));
            table.addCell(celdaMoneda(p.importe, plain, bg));
            alternate = !alternate;
        }
        doc.add(table);
    }

    private static void agregarTotales(Document doc, DatosOrden d, PdfFont bold, PdfFont plain) {
        String simbolo = simbolo(d.moneda);
        int filas = 3 + (d.mostrarDescuento ? 1 : 0) + (d.mostrarIsr ? 1 : 0);
        Table table = new Table(UnitValue.createPercentArray(new float[]{6, 2.5f, 1.5f}))
                .useAllAvailableWidth().setMarginTop(2).setKeepTogether(true);
        table.addCell(new Cell(filas, 1).setBorder(Border.NO_BORDER));
        total(table, "SUB TOTAL:", d.mostrarDescuento
                ? String.format(Locale.ROOT, simbolo + " %,.2f", d.subtotalOriginal)
                : d.subtotalTexto, bold, false);
        if (d.mostrarDescuento) {
            total(table, String.format(Locale.ROOT, "DESCUENTO (%.0f%%):", d.descuentoPorcentaje),
                    String.format(Locale.ROOT, simbolo + " -%,.2f", d.descuentoImporte), bold, false);
        }
        if (d.mostrarIsr) {
            total(table, "ISR:", String.format(Locale.ROOT, simbolo + " -%,.2f",
                    d.isrImporte), bold, false);
        }
        total(table, "IVA 16%:", d.ivaTexto, bold, false);
        total(table, "TOTAL:", d.totalTexto, bold, true);
        doc.add(table);
        doc.add(lineaSeparadora());
        if (d.observaciones != null && !d.observaciones.isBlank()) {
            doc.add(new Paragraph("OBSERVACIONES:").setFont(bold).setFontSize(8)
                    .setFontColor(AZUL).setMarginBottom(2));
            doc.add(new Paragraph(d.observaciones.trim()).setFont(plain).setFontSize(8)
                    .setMarginBottom(4));
            doc.add(lineaSeparadora());
        }
    }

    private static void agregarFactura(Document doc, DatosOrden d, PdfFont bold, PdfFont plain) {
        doc.add(new Paragraph("La generación de la factura debe incluir la siguiente información:")
                .setFont(bold).setFontSize(8).setFontColor(AZUL).setMarginBottom(2));
        Table table = new Table(UnitValue.createPercentArray(new float[]{2, 4}))
                .useAllAvailableWidth().setMarginBottom(4);
        filaVertical(table, bold, plain, "USO CFDI:", d.usoCfdi);
        filaVertical(table, bold, plain, "FORMA DE PAGO:", d.formaPago);
        filaVertical(table, bold, plain, "MÉTODO DE PAGO:", d.metodoPago);
        String moneda = d.moneda + (d.monedaDescripcion.isBlank() ? "" : " - " + d.monedaDescripcion);
        filaVertical(table, bold, plain, "TIPO DE MONEDA:", moneda);
        doc.add(table);
        doc.add(new Paragraph("Cualquier duda sobre la facturación comunicarse al departamento de contabilidad "
                + "o al correo contabilidad@rstelecom.com.mx").setFont(plain).setFontSize(8));
        doc.add(new Paragraph("El pago de facturas se realizará los días jueves y viernes, siempre y cuando sean "
                + "enviadas antes de las 12:00 hrs viéndose reflejado el pago en 24 hrs después. En caso de que "
                + "se envíen después del horario estipulado, el pago de facturas se realizará hasta la siguiente semana.")
                .setFont(plain).setFontSize(8).setMarginBottom(2));
        doc.add(lineaSeparadora());
        doc.add(new Paragraph("Orden de compra realizada en " + moneda)
                .setFont(plain).setFontSize(8).setMarginBottom(6)
                .setTextAlignment(TextAlignment.CENTER));
    }

    private static void agregarFirmas(Document doc, DatosOrden d, PdfFont bold, PdfFont plain) {
        Table table = new Table(UnitValue.createPercentArray(new float[]{1, 1}))
                .useAllAvailableWidth().setKeepTogether(true).setMarginTop(57);
        for (String[] firma : new String[][]{{"ELABORÓ", d.elaboro}, {"AUTORIZÓ", d.autorizo}}) {
            table.addCell(new Cell().add(new Paragraph("___________________________").setFont(bold)
                    .setFontSize(9).setTextAlignment(TextAlignment.CENTER))
                    .add(new Paragraph(firma[0]).setFont(bold).setFontSize(9)
                            .setTextAlignment(TextAlignment.CENTER))
                    .add(new Paragraph(firma[1]).setFont(plain).setFontSize(8)
                            .setTextAlignment(TextAlignment.CENTER)).setBorder(Border.NO_BORDER));
        }
        doc.add(table);
    }

    private static Table tablaInformacion() {
        return new Table(UnitValue.createPercentArray(new float[]{2, 4, 2, 4}))
                .useAllAvailableWidth().setMarginBottom(2);
    }

    private static void fila(Table table, PdfFont bold, PdfFont plain, String label1, String value1,
            String label2, String value2) {
        table.addCell(celdaEtiqueta(label1, bold));
        table.addCell(celdaValor(value1, plain));
        table.addCell(celdaEtiqueta(label2, bold));
        table.addCell(celdaValor(value2, plain));
    }

    private static void filaVertical(Table table, PdfFont bold, PdfFont plain, String label, String value) {
        table.addCell(celdaEtiqueta(label, bold));
        table.addCell(celdaValor(value, plain));
    }

    private static Cell celdaEtiqueta(String value, PdfFont font) {
        return new Cell().add(new Paragraph(value == null ? "" : value).setFont(font)
                .setFontSize(8).setFontColor(AZUL)).setBorder(Border.NO_BORDER).setPadding(2);
    }

    private static Cell celdaValor(String value, PdfFont font) {
        return new Cell().add(new Paragraph(value == null ? "" : value).setFont(font)
                .setFontSize(8)).setBorder(Border.NO_BORDER).setPadding(2)
                .setVerticalAlignment(VerticalAlignment.TOP);
    }

    private static Cell celda(String value, PdfFont font, TextAlignment align, DeviceRgb bg) {
        return new Cell().add(new Paragraph(value == null ? "" : value).setFont(font).setFontSize(8))
                .setTextAlignment(align).setBackgroundColor(bg).setPadding(2);
    }

    private static Cell celdaMoneda(double value, PdfFont font, DeviceRgb bg) {
        return celda(String.format(Locale.ROOT, "$%,.2f", value), font, TextAlignment.RIGHT, bg);
    }

    private static void total(Table table, String label, String value, PdfFont font, boolean destacado) {
        DeviceRgb bg = destacado ? AZUL : AZUL_CLARO;
        Cell etiqueta = new Cell().add(new Paragraph(label).setFont(font).setFontSize(destacado ? 9 : 8)
                .setFontColor(destacado ? ColorConstants.WHITE : ColorConstants.BLACK))
                .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(bg)
                .setBorder(new SolidBorder(AZUL, 1)).setPadding(3);
        Cell importe = new Cell().add(new Paragraph(value == null ? "" : value).setFont(font)
                .setFontSize(destacado ? 9 : 8)
                .setFontColor(destacado ? ColorConstants.WHITE : ColorConstants.BLACK))
                .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(bg)
                .setBorder(new SolidBorder(AZUL, 1)).setPadding(3);
        table.addCell(etiqueta);
        table.addCell(importe);
    }

    private static Table seccion(String titulo, PdfFont bold) {
        Table table = new Table(UnitValue.createPercentArray(new float[]{1})).useAllAvailableWidth();
        table.addCell(new Cell().add(new Paragraph(titulo).setFont(bold).setFontSize(8)
                .setFontColor(ColorConstants.WHITE)).setBackgroundColor(AZUL)
                .setPadding(3).setBorder(Border.NO_BORDER));
        return table;
    }

    private static Paragraph lineaSeparadora() {
        return new Paragraph("").setBorderBottom(new SolidBorder(SEPARADOR, 2)).setMarginBottom(3);
    }

    private static float calcularAltoPrimeraPagina(DatosOrden d) {
        int lineas = 0;
        for (Producto p : d.productos) {
            lineas += Math.max(1, (p.descripcion == null ? 0 : p.descripcion.length() + 39) / 40);
        }
        int observaciones = d.observaciones == null ? 0
                : Math.max(0, (d.observaciones.trim().length() + 95) / 96);
        return Math.max(PageSize.LETTER.getHeight(), 750f + d.productos.size() * 11f
                + lineas * 8f + observaciones * 9f + 45f);
    }

    private static String fecha(Date date) {
        return date == null ? "" : new SimpleDateFormat("d MMM yyyy",
                Locale.forLanguageTag("es-MX")).format(date);
    }

    private static String simbolo(String moneda) {
        if ("USD".equalsIgnoreCase(moneda)) {
            return "USD $";
        }
        if ("EUR".equalsIgnoreCase(moneda)) {
            return "€";
        }
        return "$";
    }

    private static String stripHtml(String value) {
        return value == null ? "" : value.replaceAll("<[^>]*>", "").trim();
    }
}
