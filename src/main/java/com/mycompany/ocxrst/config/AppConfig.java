package com.mycompany.ocxrst.config;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;

public final class AppConfig {

    private static final Path PROJECT_ROOT = resolveProjectRoot();
    private static final Path DATA_DIR = PROJECT_ROOT.resolve("data");
    private static final Path BASES_DIR = DATA_DIR.resolve("BASES");
    private static final Path RECURSOS_DIR = DATA_DIR.resolve("RECURSOS");

    private AppConfig() {
    }

    // El working directory al ejecutar la app puede variar (workspace raíz, subcarpeta del proyecto, etc.);
    // buscamos el proyecto real en vez de asumir que "." ya lo es.
    private static Path resolveProjectRoot() {
        Path cwd = Paths.get(".").toAbsolutePath().normalize();
        Path relativeBases = Paths.get("data", "BASES");

        Path proyectoDesdeDirectorio = buscarProyectoRoot(cwd, relativeBases);
        if (proyectoDesdeDirectorio != null) {
            return proyectoDesdeDirectorio;
        }

        try {
            Path ubicacionClases = Paths.get(AppConfig.class.getProtectionDomain()
                    .getCodeSource().getLocation().toURI()).toAbsolutePath().normalize();
            Path inicioBusqueda = ubicacionClases.toFile().isDirectory()
                    ? ubicacionClases : ubicacionClases.getParent();
            Path proyectoDesdeClases = buscarProyectoRoot(inicioBusqueda, relativeBases);
            if (proyectoDesdeClases != null) {
                return proyectoDesdeClases;
            }
        } catch (Exception ex) {
            // Si no se puede obtener la ubicación de las clases, se usa el directorio actual.
        }

        return cwd;
    }

    private static Path buscarProyectoRoot(Path inicio, Path relativeBases) {
        if (inicio == null) {
            return null;
        }
        for (Path candidate = inicio; candidate != null; candidate = candidate.getParent()) {
            if (candidate.resolve(relativeBases).toFile().isDirectory()) {
                return candidate;
            }
            File subProyecto = candidate.resolve("OrdendeComprasRST").resolve(relativeBases).toFile();
            if (subProyecto.isDirectory()) {
                return candidate.resolve("OrdendeComprasRST");
            }
        }
        return null;
    }

    public static File inventariosFile() {
        return resolveBaseFile("INVENTARIOS.xlsx");
    }

    public static File proveedoresFile() {
        return resolveBaseFile("PROVEEDORES.xlsx");
    }

    public static File productoProveedorFile() {
        return resolveBaseFile("PRODUCTOPROVEEDORES.xlsx");
    }

    public static File usuariosFile() {
        return resolveBaseFile("USUARIOS.xlsx");
    }

    public static File pagosFile() {
        return resolveBaseFile("PAGOS.xlsx");
    }

    public static File formaPagoFile() {
        return resolveBaseFile("FORMAPAGO.xlsx");
    }

    public static File metodoPagoFile() {
        return resolveBaseFile("METODOPAGO.xlsx");
    }

    public static File diasFile() {
        return resolveBaseFile("DIAS.xlsx");
    }

    public static File cfdiFile() {
        return resolveBaseFile("CFDI.xlsx");
    }

    public static File tipoMonedaFile() {
        return resolveBaseFile("TIPOMONEDA.xlsx");
    }

    public static File registroCFile() {
        return resolveBaseFile("REGISTROC.xlsx");
    }

    public static File intOrdenCompraFile() {
        return resolveBaseFile("INTORDENDECOMPRA.xlsx");
    }

    public static File actividadesFile() {
        return resolveBaseFile("ACTIVIDADES.xlsx");
    }

    public static File terminosFile() {
        return RECURSOS_DIR.resolve("Términos y Condiciones.txt").toFile();
    }

    public static File getBasesDir() {
        return BASES_DIR.toFile();
    }

    public static File resolveBaseFile(String fileName) {
        Path directPath = BASES_DIR.resolve(fileName);
        if (directPath.toFile().exists()) {
            return directPath.toFile();
        }

        File parent = directPath.getParent().toFile();
        if (parent != null && !parent.exists()) {
            parent.mkdirs();
        }
        return directPath.toFile();
    }

    public static byte[] recursoBytes(String ruta) throws IOException {
        try (InputStream entrada = AppConfig.class.getResourceAsStream(
                "/com/mycompany/ocxrst/" + ruta)) {
            if (entrada == null) {
                throw new IOException("No se encontró el recurso: " + ruta);
            }
            return entrada.readAllBytes();
        }
    }
}