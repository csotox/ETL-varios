# Pasar los eventos de Google Calendario a Google Sheets

Para ello utilizaremos Google Apps Script.

[Para entender el contexto completo, leer este artículo de LinkedIn](https://www.linkedin.com/pulse/contabilizar-el-uso-de-mi-tiempo-carlos-rafael-soto-nava-1e)

## Consideraciones

- Se requiere extraer los eventos grabados en Google Calendario y guardarlos en una hoja de Google Sheets.
- Se debe saber el ID del calendario con el que vamos a trabajar.
- El calendario debe pertenecer a la persona que va a realizar la extracción, no he probado en el escenario donde no se es el dueño del calendario.
- Cuando el código se ejecuta por primera vez, va a pedir confirmación de usuario. Por lo que debes tener las credenciales de tu cuenta de Google a la mano.

## Flujo de trabajo

### Paso # 1

- Abrir una hoja en Google Sheets y asignar un nombre.
- En el menú vamos a Herramientas -> Apps Script
- Creamos un archivo llamado getCalendarId.gs
- Agregamos el siguiente código

~~~ getCalendarId.gs

function getCalendarId() {

  const calendarList = CalendarApp.getAllOwnedCalendars()

  calendarList.forEach( cal => {
    Logger.log('Nombre calendario: %s Id: %s', cal.getName(), cal.getId())
  })

}
~~~

Sí, es una implementación de JavaScript adaptada a las aplicaciones de Google.

- Esta rutina solo trae los calendarios de los cuales somos propietarios (Se pueden traer los compartidos, pero no he probado la rutina con esos calendarios).
- Aquí podemos ver que tenemos el nombre del calendario y el ID, solo necesitamos el ID.

### Paso # 2

- Creamos un archivo llamado: exportCalendarToSheets.gs
- Agregamos el siguiente código:

~~~ exportCalendarToSheets
function exportCalendarToSheets() {

  // Reemplazamos el: **escribir_el_Id_del_calendario** por el Id del calendario
  const calendarId = "escribir_el_Id_del_calendario";
  const calendarEvent = CalendarApp.getOwnedCalendarById(calendarId);

  // Aquí tenemos el libro y la hoja donde vamos a grabar los eventos
  // importante es el valor de **reemplazar_por_el_nombre_de_la_hoja**, debe ser el nombre de una hoja que exista en el libro activo
  const hojaNombre = "reemplazar_por_el_nombre_de_la_hoja";
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = libro.getSheetByName(hojaNombre); // Nombre de la hoja donde vamos a grabar los eventos.

  // **¿Cuál es el rango de fecha que vamos a traer?**
  // El formato de Date(Año, Mes, Día)
  // Importante no quitar él -1 esto debido a que JavaScript cuenta el número de mes desde el cero
  // eso quiere decir que enero es el mes cero para js. Enero = 0
  const fInicio = new Date(2022, 1 - 1, 1);
  const fFin = new Date(2022, 12 - 1, 31);
  
  const eventCalendar = calendarEvent.getEvents(fInicio, fFin);

  eventCalendar.forEach( e => {
    const eventId = e.getId();
    const inicio = e.getStartTime();
    const fin = e.getEndTime();
    const titulo = e.getTitle();
    const color = e.getColor();
    const nota = e.getDescription();

    // Calculo el tiempo total de duración del evento, de esta manera no tengo que hacerlo
    // mas adelante
    const tiempo = (((fin-inicio) / 1000) / 60) / 60;

    // Ignoro los eventos que su color es granito.
    // pueden ser las siguientes razones:
    // 1. El evento lo asigno un tercero. No tengo control sobre si el tercero lo borra
    // 2. Eventos agendados que no deseo ignorar
    // 3. Eventos agendados con terceros que no quiero que se vea la descripción.
    if (color == 8) return;

    // Como la clasificación debe estar en la primera línea de la descripción del evento.
    // Busco el primer salto de línea
    const posSaltoX = nota.indexOf("<br>");
    var posSalto = nota.indexOf("\n");

    if (posSalto == -1 && posSaltoX > 0) {
      posSalto = posSaltoX;
    }

    // Traigo la clasificación en caso de existir.
    var clasifica = "";
    if (posSalto == -1) {
      clasifica = nota;
    } else {
      clasifica = nota.slice(0, posSalto);
    }

    console.log("Evento", inicio, "-", titulo, "-", posSalto, posSaltoX, "-", clasifica);

    // Agrego una nueva línea en el libro de Sheets con la información del evento
    hojaOrigen.appendRow([eventId, inicio, fin, titulo, color, tiempo, clasifica]);
  })

}
~~~

- Podemos notar que este código leer todos los eventos en un rango de fecha.
- Ignora los eventos que tienen un determinado color.
- Cada evento lo graba en una línea en el Sheets activo.
- Ahora puedes utilizar esta información para analizar.

## Nota: Para ejecutar el código se debe presionar el botón de Ejecutar
