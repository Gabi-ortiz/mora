/**
 * Script de cada planilla de conciliación (un banco / una cuenta).
 * Toda la lógica está en la biblioteca "Motor" (proyecto "Motor Conciliaciones"): este archivo solo arma el menú y
 * llama al motor. Las mejoras se cargan una sola vez en el motor y llegan a todas las planillas.
 * Lo que cambia de un banco a otro va en la hoja "Parámetros" de cada planilla (cuentas E y O, link de la base, etc.).
 */
function onOpen() { Motor.crearMenu(); }
function crearHojasEntrada() { Motor.crearHojasEntrada(); }
function procesarConciliacion() { Motor.procesarConciliacion(); }
function cerrarMes() { Motor.cerrarMes(); }
function revisarPendientes() { Motor.revisarPendientes(); }
function agregarReglasNuevas() { Motor.agregarReglasNuevas(); }
