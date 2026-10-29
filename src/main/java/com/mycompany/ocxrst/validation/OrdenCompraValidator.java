package com.mycompany.ocxrst.validation;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public final class OrdenCompraValidator {

    private OrdenCompraValidator() {
    }

    public static String validar(Datos datos) {
        List<String> faltantes = new ArrayList<>();

        if (datos.proveedorId().trim().isEmpty()) {
            faltantes.add("• ID PROVEEDOR");
        } else if (datos.proveedorNombre().isEmpty()
                || datos.proveedorNombre().equals("Proveedor no encontrado")
                || datos.proveedorNombre().equals("Error")) {
            faltantes.add("• PROVEEDOR (ID no encontrado, verifique el dato)");
        }

        if (datos.fecha() == null) {
            faltantes.add("• FECHA");
        }
        if (datos.cotizacion().trim().isEmpty()) {
            faltantes.add("• COTIZACIÓN");
        }
        if (!seleccionValida(datos.cfdi())) {
            faltantes.add("• USO CFDI");
        }
        if (!seleccionValida(datos.solicitante())) {
            faltantes.add("• SOLICITANTE");
        }
        if (datos.proyecto().trim().isEmpty()) {
            faltantes.add("• PROYECTO");
        }
        if (!seleccionValida(datos.pago())) {
            faltantes.add("• PAGO");
        }
        if (!seleccionValida(datos.formaPago())) {
            faltantes.add("• FORMA DE PAGO");
        }
        if (!seleccionValida(datos.metodoPago())) {
            faltantes.add("• MÉTODO DE PAGO");
        }
        if (datos.fechaEntrega() == null) {
            faltantes.add("• FECHA DE ENTREGA");
        }
        if (datos.fechaPago() == null) {
            faltantes.add("• FECHA DE PAGO");
        }
        if (datos.validarLimiteFechas() && !datos.administrador()) {
            if (esFechaAnteriorAHoy(datos.fechaEntrega())) {
                faltantes.add("• FECHA DE ENTREGA (no puede ser anterior a hoy)");
            }
            if (esFechaAnteriorAHoy(datos.fechaPago())) {
                faltantes.add("• FECHA DE PAGO (no puede ser anterior a hoy)");
            }
        }
        if (!seleccionValida(datos.moneda())) {
            faltantes.add("• TIPO DE MONEDA");
        }
        if (datos.tipoCambioVisible() && datos.tipoCambio().trim().isEmpty()) {
            faltantes.add("• TIPO DE CAMBIO");
        }
        if (!seleccionValida(datos.elaboro())) {
            faltantes.add("• ELABORÓ");
        }
        if (!seleccionValida(datos.autorizo())) {
            faltantes.add("• AUTORIZÓ");
        }
        if (!datos.tieneProductos()) {
            faltantes.add("• PRODUCTOS (la tabla debe contener al menos un producto)");
        }

        if (faltantes.isEmpty()) {
            return null;
        }
        return "Por favor, complete los siguientes campos antes de continuar:\n\n"
                + String.join("\n", faltantes);
    }

    private static boolean seleccionValida(String valor) {
        if (valor == null) {
            return false;
        }
        String texto = valor.trim();
        return !texto.isEmpty()
                && !texto.equals("** Selecciona una opción **")
                && !texto.equals("--");
    }

    private static boolean esFechaAnteriorAHoy(Date fecha) {
        if (fecha == null) {
            return false;
        }
        java.time.LocalDate fechaSeleccionada = fecha.toInstant()
                .atZone(java.time.ZoneId.systemDefault()).toLocalDate();
        return fechaSeleccionada.isBefore(java.time.LocalDate.now());
    }

    public record Datos(
            String proveedorId,
            String proveedorNombre,
            Date fecha,
            String cotizacion,
            String cfdi,
            String solicitante,
            String proyecto,
            String pago,
            String formaPago,
            String metodoPago,
            Date fechaEntrega,
            Date fechaPago,
            boolean validarLimiteFechas,
            boolean administrador,
            String moneda,
            boolean tipoCambioVisible,
            String tipoCambio,
            String elaboro,
            String autorizo,
            boolean tieneProductos) {
    }
}
