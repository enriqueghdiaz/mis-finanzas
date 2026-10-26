# Mis Finanzas: guía de instalación

Es una app web instalable (PWA). En el Redmi Note 14 se instala desde Chrome y funciona como cualquier otra app: tiene su icono y se abre a pantalla completa. Lee y escribe directamente en tu Excel de OneDrive a través de Microsoft Graph. Tus datos no pasan por ningún otro servidor.

La puesta en marcha lleva unos 20 minutos y solo se hace una vez.

---

## 1. Prepara el Excel (2 min)

1. El archivo tiene que estar en tu OneDrive (por ejemplo, `20260925 Finanzas Personales.xlsm`).
2. **Muy recomendable:** en la hoja **BDD P&G**, la columna **I** tiene un espacio escrito en ~1 millón de filas. Por eso el archivo ocupa más de 100 MB por dentro, y Excel online (y la app) van lentos o pueden fallar. Si no usas esa columna, selecciónala entera, pulsa **Supr** y guarda.
3. No cambies el nombre de las hojas `PPTO`, `BCE`, `Ahorro - Inversión` ni de la tabla `BDD_transacciones`. Si lo haces, actualiza también `config.js`.
4. Si al conectar ves un error de formato no admitido por tener macros, guarda una copia como `.xlsx`. La única macro es el botón «Guardar» de la hoja Registro, y la app ya hace ese trabajo.

## 2. Publica la app gratis en GitHub Pages (5 min)

1. Crea una cuenta en github.com, si no la tienes.
2. **New repository** → nombre `mis-finanzas` → **Public** → *Create repository*.
3. **Add file → Upload files** → arrastra **el contenido** de la carpeta `mis-finanzas` (index.html, style.css, config.js, sw.js, manifest.webmanifest y las carpetas `js` e `icons`) → *Commit changes*.
4. **Settings → Pages** → *Source: Deploy from a branch* → rama `main`, carpeta `/ (root)` → *Save*.
5. En 1–2 minutos tendrás la dirección: `https://TU-USUARIO.github.io/mis-finanzas/`

> El repositorio solo contiene el código de la app. Tus números siguen estando únicamente en tu OneDrive.

## 3. Registra la app en Microsoft (10 min)

Este paso sirve para que Microsoft deje a *tu* app abrir *tu* OneDrive.

1. Entra en **portal.azure.com** con la misma cuenta de Microsoft del OneDrive.
   - Si te dice que tu cuenta no tiene directorio o inquilino, crea una cuenta gratuita de Azure (azure.microsoft.com/free) y vuelve. El registro de apps es gratis.
2. Busca **App registrations** (Registros de aplicaciones) → **New registration**:
   - **Name:** `Mis Finanzas`
   - **Supported account types:** *Personal Microsoft accounts only* (si tu OneDrive es personal)
   - **Redirect URI:** plataforma **Single-page application (SPA)**, con la dirección del paso 2 **exacta y con la barra final**: `https://TU-USUARIO.github.io/mis-finanzas/`
   - **Register**
3. Copia el **Application (client) ID**.
4. **API permissions → Add a permission → Microsoft Graph → Delegated** → marca `Files.ReadWrite` (User.Read ya viene) → *Add permissions*.
5. Pega el Client ID en `config.js` (en GitHub: abre el archivo → ✏️ → `CLIENT_ID: "..."` → *Commit*). Otra opción es pegarlo en la pantalla de inicio de la app la primera vez.

## 4. Instálala en el Redmi Note 14 (2 min)

1. Abre la dirección en **Chrome** → **Conectar con OneDrive** → inicia sesión → acepta los permisos.
2. Elige tu Excel de la lista. Si no aparece, búscalo por nombre o escribe su ruta.
3. Menú **⋮ → Instalar aplicación** (o *Añadir a pantalla de inicio*).
   - En HyperOS/MIUI, si no aparece el icono: *Ajustes → Aplicaciones → Chrome → Otros permisos → Accesos directos en pantalla de inicio* → Permitir.
4. Si mantienes pulsado el icono, tienes el atajo **Nuevo gasto**.

¿Quieres verla antes de configurar nada? En la primera pantalla pulsa **Probar con datos de ejemplo**.

---

## Qué hace cada sección y dónde escribe en el Excel

| En la app | Lee | Escribe |
|---|---|---|
| **＋ Nuevo movimiento** (gasto, ingreso o ahorro) | tabla `BDD_transacciones` | Añade una fila arriba del todo, como tu macro. Transacción, Naturaleza y Periodicidad se rellenan solas según la categoría. El ahorro se guarda como *Gasto · Esencial · Fijo · Ahorro*. |
| **Movimientos**: editar o borrar | `BDD_transacciones` | Busca la fila por su contenido (no por posición) antes de cambiarla. |
| **Presupuesto** | `PPTO`: columnas de mes (jul, ago, sep, oct, nov, dic) | La celda de esa categoría y ese mes. «Editar mes completo» permite copiar el mes anterior y conserva las fórmulas de subtotales. |
| **Análisis** | movimientos + presupuesto | Nada, solo gráficos. |
| **Patrimonio → Balance** | `BCE` (activos C:I, pasivos L:R) | Los saldos del mes que registres. |
| **Patrimonio → Inversiones** | `BCE › Inversiones` y `Ahorro - Inversión` | `Ahorro - Inversión!C2` (importe a invertir). |
| **Patrimonio → Ahorro** | `PPTO` fila Ahorro y tasa `B17` | `PPTO!B17` (tasa de ahorro). Las metas y la proyección se guardan solo en el móvil. |

Los importes reales del mes se calculan igual que tu hoja P&G: suma por categoría según la fecha del movimiento.

## Bueno saber

- **Sin cobertura:** puedes seguir apuntando gastos. Quedan como «pendiente» y se suben solos cuando vuelve la conexión.
- **Excel y app a la vez:** puedes tener el Excel abierto en el ordenador. Toca la etiqueta de estado (arriba) o *Ajustes → Actualizar* para recargar.
- **Sesión:** Microsoft pide volver a entrar si pasas más de ~24 h sin abrir la app. Toca «Inicia sesión» en la etiqueta de estado; es un toque.
- **Presupuesto con fórmula** (p. ej. `=8.25+144.09`): si lo editas desde la app, se sustituye por el importe final. La app te avisa antes.
- **Meses nuevos (2027):** si añades columnas de mes en `PPTO` (con la fecha en la cabecera, como ahora), la app las detecta sola. `BCE` lee de junio a diciembre; para 2027 dime cómo vas a organizar el balance y adapto el rango.
- **Privacidad:** el inicio de sesión guarda los tokens solo en tu móvil. La app pide únicamente permiso para tus archivos (`Files.ReadWrite`) y tu nombre (`User.Read`). Puedes revocarlo en account.live.com/consent/Manage.

---

## Novedades v1.4

- **Partidas con cálculos:** al editar una partida puedes escribir `8,25 + 144,09 − 30` (usa la barra **+ − ( )** sobre el teclado). Se guarda en PPTO como fórmula (`=8.25+144.09-30`), igual que hacías a mano. Tus fórmulas antiguas se abren como líneas.
- **Desglose y recordatorios:** cada línea puede llevar un concepto («Seguro hogar», «Seguro moto»…) y la partida una nota. Se guardan en una hoja nueva del Excel, **App_Notas**, que la app crea sola. Las partidas con notas llevan 📝.
- **Contador «Por asignar»:** ingresos − ahorro − gastos presupuestados del mes. Si entran ingresos que no estaban previstos, el contador sube con todo el importe.
- **Planificar** (pestaña dentro de Presupuesto): consejos según tus últimos meses, referencia por partida (lo gastado en los últimos meses frente a lo presupuestado), dónde fallaba el presupuesto, fijos frente a variables, y «Preparar mes con tus medias».
- **Editar mes completo:** rellena con el presupuesto del mes anterior, con tu media real o con lo real del último mes, y ves el contador moverse mientras escribes. Las partidas con desglose se respetan.
- **Movimientos:** filtra por categoría y mes, ordena por fecha, importe o concepto, busca por concepto o importe exacto (p. ej. `46,95`), y ve el total de lo filtrado. Desde Análisis, al tocar una barra de «Real vs presupuesto» se abren sus movimientos.

## Novedades v1.5 y v1.6

- **Contador en Presupuesto:** arriba del todo ves **«Por asignar»** (ingresos − ahorro − gastos presupuestados) y, a su lado, **«Pendiente de gastar»**: lo que te queda de cada partida sumado (lo que te pasas en una no resta de las demás; se muestra aparte como exceso). «Ver detalle» despliega el cálculo. El contador se desplaza con la página y no aparece en la barra de abajo.
- **El ahorro de los ingresos extra no se separa solo.** Al registrar un ingreso que supera lo previsto, la app te pregunta: apartar tu porcentaje, otra cantidad o nada. Si apartas algo, se crea un movimiento de Ahorro con la nota «Ahorro extra», y el contador lo descuenta. Puedes desactivar la pregunta en ⚙️ Ajustes.
- **Actualizaciones:** a partir de ahora cada paquete trae todos los archivos de la app (menos `config.js`, para no borrar tu Client ID). Súbelos todos. Si falta alguno, la app te avisa con «Actualización incompleta».

## Novedades v1.7: la partida de Ahorro

- En **Presupuesto** aparece arriba la fila **Ahorro previsto** (apartado / previsto). Al tocarla puedes:
  - **Cambiar el importe** de ese mes, con sumas, restas, conceptos y nota, como cualquier partida.
  - **Volver a tu tasa** («Tu tasa: 25% de ingresos»): se guarda la fórmula original del Excel y se recalcula sola.
  - **Retirar del ahorro para gastar:** indica cuánto y el motivo. Se resta del ahorro previsto (p. ej. `=507.5-50`), así sube «Por asignar», y opcionalmente se registra la retirada como movimiento de Ahorro en negativo.
- Si ese mes ya has traspasado el ahorro, bajar el previsto no libera dinero: para eso está «Retirar del ahorro».
- En **Editar mes completo** también puedes cambiar el ahorro previsto, y el contador se mueve mientras escribes.
- En **Nuevo movimiento → Ahorro** puedes elegir **Aportar** o **Retirar**. Las retiradas se ven como «← importe».

## Novedades v1.8: cuentas y fotos del presupuesto

**Cuentas (Openbank y Edenred)**
- Al registrar un gasto de **Restaurantes** (o el ingreso de **Dietas**) eliges si es con **Openbank** o con **Edenred**. En el resto de categorías solo existe Openbank. La app recuerda la última cuenta usada en cada categoría.
- Se guarda en una columna nueva, **«Cuenta»**, de tu tabla de movimientos (BDD_transacciones), que la app crea sola. Los movimientos antiguos con «Edenred» en la nota se reconocen automáticamente.
- En la partida de **Restaurantes** indicas **«De ello con Edenred»**, y en **Dietas** cuánto entra en Edenred. Esa parte no cuenta como dinero reasignable: el contador muestra aparte «🎫 Edenred: X € por asignar (solo Restaurantes)», y «Pendiente de gastar» dice cuánto es de Edenred.
- En Movimientos puedes filtrar por cuenta.
- Para añadir más cuentas en el futuro: `ACCOUNTS` en `config.js`.

**Congelar presupuesto (fotos)**
- En Presupuesto, **📸 Congelar** guarda una foto del presupuesto del mes tal como está, con el nombre que quieras (Inicial, Reajuste 15 oct…). Antes de guardarla te enseña qué ha cambiado desde la última.
- La primera vez que cambias el presupuesto de un mes, la app guarda sola una foto «Inicial (automática)». Se puede desactivar en ⚙️ Ajustes.
- Las fotos se guardan en una hoja nueva del Excel, **«App_Fotos»**, con una fila por partida (fecha, mes, etiqueta, categoría, importe), lista para tablas dinámicas.
- **Historial del mes:** cada foto con lo que subiste o bajaste respecto a la anterior.
- **En Planificar:**
  - «Reorganizaciones del presupuesto»: qué partidas cambias más, cuánto subes y bajas, y cuánto cambia cada partida de media del inicio al final del mes.
  - «¿Dónde fallaba el presupuesto?» se puede ver frente al **presupuesto inicial** o al final.
  - Consejos del tipo «Restaurantes: la subes durante el mes; empieza con ~X €».
- **En Análisis:** «Real vs presupuesto» marca con una raya gris el presupuesto que tenías al empezar el mes.
