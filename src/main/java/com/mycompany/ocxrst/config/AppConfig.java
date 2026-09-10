package com.mycompany.ocxrst.config;

import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;

public final class AppConfig {

    private static final Path PROJECT_ROOT = resolveProjectRoot();
    private static final Path BASES_DIR = PROJECT_ROOT.resolve("src").resolve("main").resolve("java")
            .resolve("com").resolve("mycompany").resolve("ocxrst").resolve("BASES");

    private AppConfig() {
    }

    // El working directory al ejecutar la app puede variar (workspace raíz, subcarpeta del proyecto, etc.);
    // buscamos el proyecto real en vez de asumir que "." ya lo es.
    private static Path resolveProjectRoot() {
        Path cwd = Paths.get(".").toAbsolutePath().normalize();
        Path relativeBases = Paths.get("src", "main", "java", "com", "mycompany", "ocxrst", "BASES");

        for (Path candidate = cwd; candidate != null; candidate = candidate.getParent()) {
            if (candidate.resolve(relativeBases).toFile().isDirectory()) {
                return candidate;
            }
            File subProyecto = candidate.resolve("OrdendeComprasRST").resolve(relativeBases).toFile();
            if (subProyecto.isDirectory()) {
                return candidate.resolve("OrdendeComprasRST");
            }
        }

        return cwd;
    }

    public static File inventariosFile() {
        return resolveBaseFile("INVENTARIOS.xlsx");
    }

    public static File proveedoresFile() {
        return resolveBaseFile("PROVEEDORES.xlsx");
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

    public static File terminosFile() {
        return PROJECT_ROOT.resolve("src").resolve("main").resolve("java")
                .resolve("com").resolve("mycompany").resolve("ocxrst")
                .resolve("RECURSOS").resolve("Términos y Condiciones.txt").toFile();
    }

    public static File getBasesDir() {
        return BASES_DIR.toFile();
    }

    public static File resolveBaseFile(String fileName) {
        Path directPath = BASES_DIR.resolve(fileName);
        if (directPath.toFile().exists()) {
            return directPath.toFile();
        }

        Path fallback = PROJECT_ROOT.resolve("src").resolve("main").resolve("java")
                .resolve("com").resolve("mycompany").resolve("ocxrst").resolve("BASES")
                .resolve(fileName);
        if (fallback.toFile().exists()) {
            return fallback.toFile();
        }

        File parent = directPath.getParent().toFile();
        if (parent != null && !parent.exists()) {
            parent.mkdirs();
        }
        return directPath.toFile();
    }
}
