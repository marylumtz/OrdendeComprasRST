package com.mycompany.ocxrst.repository;

import com.mycompany.ocxrst.config.AppConfig;
import com.mycompany.ocxrst.model.Tarea;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class TareaRepository {
    public static final String[] ENCABEZADOS = {"FOLIO", "DESCRIPCI\u00d3N",
        "TIPO", "ESTADO", "PRIORIDAD", "RESPONSABLE", "FECHA CREACI\u00d3N",
        "FECHA VENCIMIENTO", "ACCIONES"};
    private static final String ENCABEZADO_ANTERIOR = "T\u00cdTULO / DESCRIPCI\u00d3N";
    private static final String[] DETALLES = {"FOLIO", "DESCRIPCI\u00d3N DETALLADA",
        "FECHA INICIO", "FOLIO RELACIONADO", "OBSERVACIONES"};
    private static final Object LOCK = new Object();
    private final Path archivo;

    public TareaRepository() {
        this(AppConfig.resolveBaseFile("TAREAS.xlsx").toPath());
    }

    public TareaRepository(Path archivo) {
        this.archivo = archivo.toAbsolutePath();
    }

    public List<Tarea> listar() throws IOException {
        synchronized (LOCK) {
            if (!Files.exists(archivo)) {
                return List.of();
            }
            try (Workbook libro = abrir()) {
                return leer(libro);
            }
        }
    }

    public Tarea guardar(Tarea tarea) throws IOException {
        tarea.validar();
        synchronized (LOCK) {
            try (Workbook libro = abrir()) {
                List<Tarea> tareas = new ArrayList<>(leer(libro));
                String folio = tarea.folio();
                int indice = -1;
                if (folio != null && !folio.isBlank()) {
                    for (int i = 0; i < tareas.size(); i++) {
                        if (tareas.get(i).folio().equals(folio)) {
                            indice = i;
                            break;
                        }
                    }
                    if (indice < 0) {
                        throw new IOException("La tarea ya no existe. Actualiza la mesa de trabajo.");
                    }
                } else {
                    int siguiente = 1;
                    for (Tarea existente : tareas) {
                        siguiente = Math.max(siguiente, Integer.parseInt(existente.folio().substring(2)) + 1);
                    }
                    Sheet control = libro.getSheet("CONTROL");
                    if (control != null && control.getRow(1) != null) {
                        siguiente = Math.max(siguiente, Integer.parseInt(celda(control.getRow(1), 0)));
                    }
                    folio = String.format(java.util.Locale.ROOT, "T-%04d", siguiente);
                    if (control == null) {
                        control = libro.createSheet("CONTROL");
                        control.createRow(0).createCell(0).setCellValue("SIGUIENTE FOLIO");
                    }
                    control.createRow(1).createCell(0).setCellValue(String.valueOf(siguiente + 1));
                }
                List<String> adjuntos = copiarAdjuntos(tarea.adjuntos(), folio);
                Tarea guardada = new Tarea(folio, tarea.titulo().trim(), tarea.tipo(), tarea.estado(),
                        tarea.prioridad(), tarea.responsable(), indice < 0 ? tarea.creacion() : tareas.get(indice).creacion(),
                        tarea.vencimiento(), tarea.descripcion(), tarea.inicio(), tarea.folioRelacionado(),
                        tarea.observaciones(), adjuntos);
                if (indice < 0) {
                    tareas.add(guardada);
                } else {
                    tareas.set(indice, guardada);
                }
                escribir(libro, tareas);
                guardarLibro(libro, archivo);
                return guardada;
            } catch (IllegalArgumentException ex) {
                throw new IOException("El archivo de tareas contiene datos inv\u00e1lidos.", ex);
            }
        }
    }

    public void desactivar(String folio) throws IOException {
        synchronized (LOCK) {
            try (Workbook libro = abrir()) {
                List<Tarea> tareas = new ArrayList<>(leer(libro));
                int indice = -1;
                for (int i = 0; i < tareas.size(); i++) {
                    if (tareas.get(i).folio().equals(folio)) {
                        indice = i;
                        break;
                    }
                }
                if (indice < 0) {
                    throw new IOException("La tarea ya no existe. Actualiza la mesa de trabajo.");
                }
                Tarea actual = tareas.get(indice);
                tareas.set(indice, new Tarea(actual.folio(), actual.titulo(), actual.tipo(), "Cancelada",
                        actual.prioridad(), actual.responsable(), actual.creacion(), actual.vencimiento(),
                        actual.descripcion(), actual.inicio(), actual.folioRelacionado(), actual.observaciones(),
                        actual.adjuntos()));
                escribir(libro, tareas);
                guardarLibro(libro, archivo);
            }
        }
    }

    public void exportar(List<Tarea> tareas, Path destino) throws IOException {
        if (destino.toAbsolutePath().equals(archivo)) {
            throw new IOException("Selecciona un archivo distinto al registro TAREAS.xlsx.");
        }
        try (Workbook libro = new XSSFWorkbook()) {
            escribir(libro, tareas);
            guardarLibro(libro, destino.toAbsolutePath());
        }
    }

    private Workbook abrir() throws IOException {
        if (!Files.exists(archivo)) {
            return new XSSFWorkbook();
        }
        try (InputStream entrada = Files.newInputStream(archivo)) {
            return new XSSFWorkbook(entrada);
        }
    }

    private List<Tarea> leer(Workbook libro) throws IOException {
        Sheet hoja = libro.getSheet("TAREAS");
        if (hoja == null) {
            if (libro.getNumberOfSheets() != 0) {
                throw new IOException("No se encontr\u00f3 la hoja TAREAS.");
            }
            return List.of();
        }
        for (int i = 0; i < ENCABEZADOS.length; i++) {
            String encabezado = celda(hoja.getRow(0), i);
            boolean esEncabezadoAnterior = i == 1 && ENCABEZADO_ANTERIOR.equals(encabezado);
            if (!ENCABEZADOS[i].equals(encabezado) && !esEncabezadoAnterior) {
                throw new IOException("La estructura de TAREAS.xlsx no coincide con las columnas esperadas.");
            }
        }
        List<Tarea> tareas = new ArrayList<>();
        Set<String> folios = new HashSet<>();
        for (int i = 1; i <= hoja.getLastRowNum(); i++) {
            Row fila = hoja.getRow(i);
            String folio = celda(fila, 0);
            if (fila == null || folio.isBlank()) {
                continue;
            }
            try {
                if (!folio.matches("T-\\d{4,}") || !folios.add(folio)) {
                    throw new IllegalArgumentException("Folio inv\u00e1lido o duplicado: " + folio);
                }
                Row detalles = buscar(libro.getSheet("DETALLES"), folio);
                if (detalles == null) {
                    throw new IllegalArgumentException("Faltan los detalles de " + folio);
                }
                List<String> adjuntos = new ArrayList<>();
                Sheet archivos = libro.getSheet("ADJUNTOS");
                if (archivos != null) {
                    for (Row adjunto : archivos) {
                        if (folio.equals(celda(adjunto, 0))) {
                            adjuntos.add(celda(adjunto, 1));
                        }
                    }
                }
                Tarea tarea = new Tarea(folio, celda(fila, 1), celda(fila, 2), celda(fila, 3),
                        celda(fila, 4), celda(fila, 5), fecha(fila, 6), fecha(fila, 7),
                        celda(detalles, 1), fecha(detalles, 2), celda(detalles, 3), celda(detalles, 4), adjuntos);
                tarea.validar();
                tareas.add(tarea);
            } catch (IllegalArgumentException | java.time.DateTimeException ex) {
                throw new IOException("Revisa la fila " + (i + 1) + " de TAREAS.xlsx: " + ex.getMessage(), ex);
            }
        }
        return tareas;
    }

    private static LocalDate fecha(Row fila, int columna) {
        if (fila != null && fila.getCell(columna) != null
                && fila.getCell(columna).getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC
                && org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(fila.getCell(columna))) {
            return fila.getCell(columna).getLocalDateTimeCellValue().toLocalDate();
        }
        return LocalDate.parse(celda(fila, columna));
    }

    private static String celda(Row fila, int columna) {
        return fila == null ? "" : new DataFormatter().formatCellValue(fila.getCell(columna)).trim();
    }

    private static Row buscar(Sheet hoja, String folio) {
        if (hoja != null) {
            for (Row fila : hoja) {
                if (folio.equals(celda(fila, 0))) {
                    return fila;
                }
            }
        }
        return null;
    }

    private static Sheet prepararHoja(Workbook libro, String nombre, String[] encabezados) {
        Sheet hoja = libro.getSheet(nombre);
        if (hoja == null) {
            hoja = libro.createSheet(nombre);
        }
        for (int i = hoja.getLastRowNum(); i >= 0; i--) {
            Row fila = hoja.getRow(i);
            if (fila != null) {
                hoja.removeRow(fila);
            }
        }
        escribirFila(hoja.createRow(0), encabezados);
        hoja.createFreezePane(0, 1);
        return hoja;
    }

    private static void escribirFila(Row fila, String... valores) {
        for (int i = 0; i < valores.length; i++) {
            fila.createCell(i).setCellValue(valores[i] == null ? "" : valores[i]);
        }
    }

    private static void escribir(Workbook libro, List<Tarea> tareas) {
        Sheet principal = prepararHoja(libro, "TAREAS", ENCABEZADOS);
        Sheet detalles = prepararHoja(libro, "DETALLES", DETALLES);
        Sheet adjuntos = prepararHoja(libro, "ADJUNTOS", new String[]{"FOLIO", "ARCHIVO"});
        int fila = 1;
        int archivo = 1;
        for (Tarea tarea : tareas) {
            escribirFila(principal.createRow(fila), tarea.folio(), tarea.titulo(), tarea.tipo(), tarea.estado(),
                    tarea.prioridad(), tarea.responsable(), tarea.creacion().toString(), tarea.vencimiento().toString(),
                    "VER | EDITAR | DESACTIVAR");
            escribirFila(detalles.createRow(fila++), tarea.folio(), tarea.descripcion(), tarea.inicio().toString(),
                    tarea.folioRelacionado(), tarea.observaciones());
            for (String ruta : tarea.adjuntos()) {
                escribirFila(adjuntos.createRow(archivo++), tarea.folio(), ruta);
            }
        }
        for (int i = 0; i < ENCABEZADOS.length; i++) {
            principal.setColumnWidth(i, (i == 1 ? 45 : 22) * 256);
        }
        libro.setSheetVisibility(libro.getSheetIndex(detalles), SheetVisibility.VERY_HIDDEN);
        libro.setSheetVisibility(libro.getSheetIndex(adjuntos), SheetVisibility.VERY_HIDDEN);
        Sheet control = libro.getSheet("CONTROL");
        if (control != null) {
            libro.setSheetVisibility(libro.getSheetIndex(control), SheetVisibility.VERY_HIDDEN);
        }
        libro.setActiveSheet(libro.getSheetIndex(principal));
    }

    private List<String> copiarAdjuntos(List<String> rutas, String folio) throws IOException {
        List<String> resultado = new ArrayList<>();
        Path carpeta = archivo.getParent().resolve("ADJUNTOS_TAREAS").resolve(folio);
        for (String ruta : rutas) {
            Path origen = Path.of(ruta).toAbsolutePath().normalize();
            if (!Files.isRegularFile(origen)) {
                throw new IOException("No se encontr\u00f3 el adjunto: " + origen.getFileName());
            }
            String nombre = origen.getFileName().toString();
            String extension = nombre.substring(nombre.lastIndexOf('.') + 1).toLowerCase(java.util.Locale.ROOT);
            if (!List.of("pdf", "jpg", "jpeg", "png", "xlsx").contains(extension)
                    || Files.size(origen) > 10L * 1024 * 1024) {
                throw new IOException("Adjunto no permitido: " + nombre + ". Usa PDF, JPG, PNG o XLSX (m\u00e1x. 10 MB).");
            }
            if (!origen.getParent().equals(carpeta)) {
                Files.createDirectories(carpeta);
                Path destino = carpeta.resolve(java.util.UUID.randomUUID() + "_" + nombre);
                Files.copy(origen, destino);
                origen = destino;
            }
            resultado.add(origen.toString());
        }
        return resultado;
    }

    private static void guardarLibro(Workbook libro, Path destino) throws IOException {
        Files.createDirectories(destino.getParent());
        Path temporal = Files.createTempFile(destino.getParent(), "tareas-", ".xlsx");
        try {
            try (OutputStream salida = Files.newOutputStream(temporal)) {
                libro.write(salida);
            }
            try {
                Files.move(temporal, destino, StandardCopyOption.ATOMIC_MOVE, StandardCopyOption.REPLACE_EXISTING);
            } catch (AtomicMoveNotSupportedException ex) {
                Files.move(temporal, destino, StandardCopyOption.REPLACE_EXISTING);
            }
        } finally {
            Files.deleteIfExists(temporal);
        }
    }
}
