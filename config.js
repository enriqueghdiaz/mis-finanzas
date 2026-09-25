// ─────────────────────────────────────────────────────────────
// Configuración de la app "Mis Finanzas"
// Solo tienes que cambiar CLIENT_ID (ver GUIA_INSTALACION).
// También puedes ponerlo desde Ajustes dentro de la app.
// ─────────────────────────────────────────────────────────────
window.APP_CONFIG = {
  // "Id. de aplicación (cliente)" que te da Microsoft al registrar la app
  CLIENT_ID: "",

  // "consumers" = cuentas personales de Microsoft (Outlook/Hotmail/Live o Gmail registrado en Microsoft).
  // Usa "common" si tu OneDrive es de trabajo/escuela.
  AUTHORITY: "consumers",

  // Nombre (o parte del nombre) del Excel en OneDrive, para buscarlo la primera vez
  FILE_SEARCH: "Finanzas Personales",

  // Nombres de hojas y tabla del Excel (no cambiar salvo que los renombres)
  SHEETS: {
    budget: "PPTO",
    balance: "BCE",
    invest: "Ahorro - Inversión",
    txTable: "BDD_transacciones"
  }
};
