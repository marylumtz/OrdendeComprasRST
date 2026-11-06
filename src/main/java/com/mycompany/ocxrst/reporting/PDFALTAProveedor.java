package com.mycompany.ocxrst.reporting;

import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import com.itextpdf.layout.properties.VerticalAlignment;
import com.mycompany.ocxrst.config.AppConfig;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

/**
 * Genera la solicitud de alta de proveedor con el mismo estilo visual de los
 * reportes de órdenes de compra.
 */
public final class PDFALTAProveedor {

    private static final DeviceRgb AZUL = new DeviceRgb(0, 95, 131);
    private static final DeviceRgb GRIS = new DeviceRgb(90, 90, 90);
    private static final DeviceRgb BLANCO = new DeviceRgb(255, 255, 255);
    private static final SolidBorder BORDE = new SolidBorder(new DeviceRgb(0, 0, 0), 0.7f);

    private PDFALTAProveedor() {
    }

    public static final class CuentaBancaria {
        public String banco = "";
        public String numeroCuenta = "";
        public String clabe = "";
    }

    public static final class DatosProveedor {
        public String id = "";
        public String fecha = "";
        public String documento = "FO-SG-01-00";
        public String solicitante = "";
        public String departamento = "";
        public String nombre = "";
        public String rfc = "";
        public String telefono = "";
        public String correo = "";
        public String direccion = "";
        public String tipoProveedor = "";
        public String proyecto = "";
        public String observaciones = "";
        public String autorizadoPor = "";
        public final List<CuentaBancaria> cuentas = new ArrayList<>();
    }

    public static File generar(File destino, DatosProveedor datos) throws IOException {
        if (destino == null || datos == null) {
            throw new IllegalArgumentException("El destino y los datos del proveedor son obligatorios.");
        }
        File parent = destino.getParentFile();
        if (parent != null) {
            Files.createDirectories(parent.toPath());
        }

        try (PdfWriter writer = new PdfWriter(destino);
                PdfDocument pdf = new PdfDocument(writer);
                Document doc = new Document(pdf, new PageSize(PageSize.LETTER.getWidth(),
                        calcularAltoPagina(datos)))) {
            PdfFont bold = PdfFontFactory.createFont(StandardFonts.HELVETICA_BOLD);
            PdfFont plain = PdfFontFactory.createFont(StandardFonts.HELVETICA);
            doc.setMargins(28, 36, 24, 36);

            agregarEncabezado(doc, datos, bold, plain);
            doc.add(seccion("01 · DATOS DE LA SOLICITUD", bold));
            doc.add(tablaCampos(bold, plain,
                    campo("NOMBRE DEL SOLICITANTE", datos.solicitante),
                    campo("DEPARTAMENTO", datos.departamento),
                    campoCompleto("FECHA DE SOLICITUD", datos.fecha)));

            doc.add(seccion("02 · DATOS GENERALES DEL PROVEEDOR", bold));
            doc.add(tablaCampos(bold, plain,
                    campo("NOMBRE O RAZÓN SOCIAL", datos.nombre),
                    campo("RFC", datos.rfc),
                    campo("TELÉFONO", datos.telefono),
                    campo("CORREO ELECTRÓNICO", datos.correo),
                    campoCompleto("DIRECCIÓN", datos.direccion)));

            doc.add(seccion("03 · INFORMACIÓN COMERCIAL", bold));
            doc.add(tablaCampos(bold, plain,
                    campo("PROVEEDOR DE / BIENES O SERVICIOS", datos.tipoProveedor),
                    campo("PROYECTO RELACIONADO", datos.proyecto)));

            doc.add(seccion("04 · DATOS BANCARIOS", bold));
            {
                Table cuentas = new Table(UnitValue.createPercentArray(new float[]{30, 30, 40}))
                        .useAllAvailableWidth();
                encabezado(cuentas, "BANCO", bold);
                encabezado(cuentas, "NO. CUENTA", bold);
                encabezado(cuentas, "CLABE", bold);
                for (CuentaBancaria cuenta : datos.cuentas) {
                    if (cuenta == null || (vacio(cuenta.banco) && vacio(cuenta.numeroCuenta)
                            && vacio(cuenta.clabe))) {
                        continue;
                    }
                    celda(cuentas, cuenta.banco, plain);
                    celda(cuentas, cuenta.numeroCuenta, plain);
                    celda(cuentas, cuenta.clabe, plain);
                }
                if (cuentas.getNumberOfRows() == 1) {
                    celda(cuentas, "", plain);
                    celda(cuentas, "", plain);
                    celda(cuentas, "", plain);
                }
                doc.add(cuentas.setMarginBottom(8));
            }

            doc.add(seccion("05 · OBSERVACIONES ESPECÍFICAS", bold));
            doc.add(new Cell().add(new Paragraph(valor(datos.observaciones)).setFont(plain).setFontSize(9))
                    .setBorder(BORDE).setMinHeight(58).setPadding(7).setMarginBottom(8));

            Table firmas = new Table(UnitValue.createPercentArray(new float[]{50, 50}))
                    .useAllAvailableWidth().setMarginTop(16);
            firmas.addCell(firma("SOLICITADO POR", datos.solicitante, bold, plain));
            firmas.addCell(firma("AUTORIZADO POR", datos.autorizadoPor, bold, plain));
            doc.add(firmas);
            Table pie = new Table(UnitValue.createPercentArray(new float[]{70, 30}))
                    .useAllAvailableWidth().setMarginTop(10).setBorderTop(BORDE);
            pie.addCell(new Cell().add(new Paragraph("RST · RED SINERGIA DE TELECOMUNICACIONES")
                    .setFont(bold).setFontSize(7).setFontColor(GRIS))
                    .setBorder(Border.NO_BORDER).setPaddingTop(6));
            pie.addCell(new Cell().add(new Paragraph(valor(datos.documento))
                    .setFont(plain).setFontSize(7).setFontColor(GRIS)
                    .setTextAlignment(TextAlignment.RIGHT))
                    .setBorder(Border.NO_BORDER).setPaddingTop(6));
            doc.add(pie);
        }
        return destino;
    }

    private static void agregarEncabezado(Document doc, DatosProveedor datos,
            PdfFont bold, PdfFont plain) throws IOException {
        Table encabezado = new Table(UnitValue.createPercentArray(new float[]{25, 50, 25}))
                .useAllAvailableWidth().setBorder(BORDE);
        Image logo = new Image(ImageDataFactory.create(
                AppConfig.recursoBytes("IMAGENES/Logo PDF.png"))).scaleToFit(125, 48);
        encabezado.addCell(new Cell().add(logo).setBorder(Border.NO_BORDER).setPadding(5)
                .setVerticalAlignment(VerticalAlignment.MIDDLE));
        encabezado.addCell(new Cell().add(new Paragraph("SOLICITUD DE ALTA DE PROVEEDOR")
                .setFont(bold).setFontSize(13).setFontColor(AZUL)
                .setTextAlignment(TextAlignment.CENTER))
                .setBorder(Border.NO_BORDER).setPadding(5)
                .setVerticalAlignment(VerticalAlignment.MIDDLE));
        encabezado.addCell(new Cell().add(new Paragraph(valor(datos.documento))
                .setFont(bold).setFontSize(8).setTextAlignment(TextAlignment.RIGHT))
                .add(new Paragraph("ID: " + valor(datos.id)).setFont(plain).setFontSize(8)
                        .setTextAlignment(TextAlignment.RIGHT))
                .setBorder(Border.NO_BORDER).setPadding(5));
        doc.add(encabezado.setMarginBottom(4));
    }

    private static Table seccion(String texto, PdfFont bold) {
        Cell celda = new Cell().add(new Paragraph(texto).setFont(bold).setFontSize(9)
                .setFontColor(AZUL)).setBackgroundColor(new DeviceRgb(231, 242, 247))
                .setBorder(Border.NO_BORDER).setBorderLeft(new SolidBorder(new DeviceRgb(25, 157, 170), 3))
                .setPaddingLeft(8).setPaddingTop(5).setPaddingBottom(5);
        return new Table(UnitValue.createPercentArray(new float[]{100}))
                .useAllAvailableWidth().addCell(celda).setMarginTop(7).setMarginBottom(2);
    }

    private static Table tablaCampos(PdfFont bold, PdfFont plain, Cell... campos) {
        Table tabla = new Table(UnitValue.createPercentArray(new float[]{50, 50}))
                .useAllAvailableWidth();
        for (Cell campo : campos) {
            tabla.addCell(campo);
        }
        return tabla.setMarginBottom(2);
    }

    private static Cell campo(String etiqueta, String valor) {
        return campo(etiqueta, valor, 1);
    }

    private static Cell campo(String etiqueta, String valor, int columnas) {
        return new Cell(1, columnas).add(new Paragraph(etiqueta).setFontSize(7).setBold())
                .add(new Paragraph(valor(valor)).setFontSize(9).setBorderBottom(BORDE))
                .setBorder(Border.NO_BORDER).setPadding(5);
    }

    private static Cell campoCompleto(String etiqueta, String valor) {
        return campo(etiqueta, valor, 2);
    }

    private static void encabezado(Table tabla, String texto, PdfFont fuente) {
        tabla.addCell(new Cell().add(new Paragraph(texto).setFont(fuente).setFontSize(7)
                .setFontColor(BLANCO).setTextAlignment(TextAlignment.CENTER))
                .setBackgroundColor(AZUL).setBorder(BORDE).setPadding(4));
    }

    private static void celda(Table tabla, String texto, PdfFont fuente) {
        tabla.addCell(new Cell().add(new Paragraph(valor(texto)).setFont(fuente).setFontSize(8)
                .setTextAlignment(TextAlignment.CENTER)).setBorder(BORDE).setPadding(4));
    }

    private static Cell firma(String titulo, String nombre, PdfFont bold, PdfFont plain) {
        return new Cell().add(new Paragraph(titulo).setFont(bold).setFontSize(8)
                .setTextAlignment(TextAlignment.CENTER))
                .add(new Paragraph(" ").setMinHeight(56.7f).setMargin(0))
                .add(new Paragraph("____________________________\n" + valor(nombre))
                .setFont(plain).setFontSize(8).setTextAlignment(TextAlignment.CENTER)
                .setMargin(0))
                .setBorder(Border.NO_BORDER);
    }

    private static float calcularAltoPagina(DatosProveedor datos) {
        int cuentas = datos.cuentas == null ? 0 : datos.cuentas.size();
        int observaciones = datos.observaciones == null
                ? 0 : Math.max(0, (datos.observaciones.trim().length() + 95) / 96);
        return Math.max(920f, 880f + Math.max(0, cuentas - 1) * 24f
                + observaciones * 9f);
    }

    private static String valor(String texto) {
        return texto == null || texto.isBlank() ? " " : texto.trim();
    }

    private static boolean vacio(String texto) {
        return texto == null || texto.isBlank();
    }
}
