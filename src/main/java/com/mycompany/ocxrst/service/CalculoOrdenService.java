package com.mycompany.ocxrst.service;

import java.util.List;

public final class CalculoOrdenService {

    private static final double IVA = 0.16;

    private CalculoOrdenService() {
    }

    public static Resultado calcular(List<Item> items, String moneda, double tipoCambio,
            boolean aplicarDescuento, double porcentajeDescuento,
            boolean aplicarIsr, double porcentajeIsr, boolean aplicarIva) {
        double base = 0;
        for (Item item : items) {
            base += item.importe();
        }

        if (requiereConversion(moneda) && tipoCambio > 0) {
            base /= tipoCambio;
        }

        double descuento = porcentajeValido(aplicarDescuento, porcentajeDescuento)
                ? base * porcentajeDescuento / 100.0 : 0;
        double subtotal = base - descuento;

        double isr = porcentajeValido(aplicarIsr, porcentajeIsr)
                ? subtotal * porcentajeIsr / 100.0 : 0;
        subtotal -= isr;

        double iva = aplicarIva ? subtotal * IVA : 0;
        double total = subtotal + iva;

        return new Resultado(base, descuento, isr, subtotal, iva, total,
                aplicarDescuento && descuento > 0 ? porcentajeDescuento : 0,
                aplicarIsr && isr > 0 ? porcentajeIsr : 0);
    }

    public static int parseCantidad(Object valor) {
        if (valor == null) {
            return 1;
        }
        if (valor instanceof Number numero) {
            return Math.max(1, numero.intValue());
        }
        try {
            return Math.max(1, Integer.parseInt(valor.toString().trim()));
        } catch (NumberFormatException ex) {
            return 1;
        }
    }

    public static double parseNumero(Object valor) {
        if (valor == null) {
            return 0;
        }
        if (valor instanceof Number numero) {
            return numero.doubleValue();
        }
        try {
            return Double.parseDouble(valor.toString().trim().replace(",", ""));
        } catch (NumberFormatException ex) {
            return 0;
        }
    }

    private static boolean requiereConversion(String moneda) {
        return moneda != null
                && (moneda.equalsIgnoreCase("USD")
                || moneda.equalsIgnoreCase("EUR")
                || moneda.equalsIgnoreCase("CAD"));
    }

    private static boolean porcentajeValido(boolean habilitado, double porcentaje) {
        return habilitado && porcentaje > 0 && porcentaje <= 100;
    }

    public record Item(int cantidad, double precio) {
        public double importe() {
            return cantidad * precio;
        }
    }

    public record Resultado(double subtotalOriginal, double descuento, double isr,
            double subtotal, double iva, double total,
            double porcentajeDescuento, double porcentajeIsr) {
    }
}
