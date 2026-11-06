package com.mycompany.ocxrst;

import com.mycompany.ocxrst.reporting.PDFALTAProveedor;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import java.nio.file.Files;
import java.nio.file.Path;
import static org.junit.jupiter.api.Assertions.assertEquals;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.assertTrue;

class PDFALTAProveedorTest {

    @Test
    void generaSolicitudCompleta() throws Exception {
        Path destino = Files.createTempFile("solicitud-alta-", ".pdf");
        PDFALTAProveedor.DatosProveedor datos = new PDFALTAProveedor.DatosProveedor();
        datos.id = "prue03";
        datos.fecha = "02/11/2026";
        datos.solicitante = "Solicitante";
        datos.departamento = "Finanzas";
        datos.nombre = "Proveedor";
        datos.rfc = "RFC123";
        datos.telefono = "5555555555";
        datos.correo = "correo@ejemplo.com";
        datos.direccion = "Direccion";
        datos.tipoProveedor = "Bienes y servicios";
        datos.proyecto = "Proyecto";
        datos.autorizadoPor = "Autorizador";
        datos.observaciones = "Observaciones";

        PDFALTAProveedor.generar(destino.toFile(), datos);

        assertTrue(Files.size(destino) > 10_000);
        try (PdfDocument pdf = new PdfDocument(new PdfReader(destino.toFile()))) {
            assertEquals(1, pdf.getNumberOfPages());
        }
    }
}
