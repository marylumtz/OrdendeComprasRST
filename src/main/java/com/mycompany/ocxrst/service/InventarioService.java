package com.mycompany.ocxrst.service;

public final class InventarioService {

    private InventarioService() {
    }

    public static String normalizarEstado(String valor) {
        String limpio = valor == null ? "" : valor.trim();
        if (limpio.isEmpty()) {
            return "1";
        }
        if ("1".equals(limpio)) {
            return "1";
        }
        if ("0".equals(limpio)) {
            return "0";
        }
        try {
            return Double.parseDouble(limpio) > 0 ? "1" : "0";
        } catch (NumberFormatException ex) {
            if ("ACTIVO".equalsIgnoreCase(limpio)
                    || "SI".equalsIgnoreCase(limpio)
                    || "TRUE".equalsIgnoreCase(limpio)) {
                return "1";
            }
            return "0";
        }
    }
}
