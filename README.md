# Alerta personal de resoluciones para Excel Online

Este paquete contiene el complemento web para Excel Online. Lee únicamente las columnas B, E, F y G de las hojas cuyo nombre corresponda a un mes y año en español. No escribe en el libro ni envía su contenido a GitHub.

## Funcionamiento

- Escucha los cambios del libro y primero comprueba el nombre de la hoja modificada. Solo procesa hojas con nombre de mes y año en español y encabezados esperados en B1, E1, F1 y G1.
- Al cambiar una celda de F, compara la cédula de E con el resto de las filas de esa misma hoja.
- Solo alerta cuando encuentra la misma cédula con la resolución opuesta (`APLICA` frente a `NO APLICA`) en otro usuario. No incluye las propias filas de quien está usando el complemento.
- Muestra el nombre de B, la resolución y el motivo de G de cada antecedente contradictorio. Si no hay contradicción, no abre una ventana.
- El botón **Aceptar** cierra la ventana. El interruptor del panel activa o desactiva las alertas.
- Solo muestra alertas cuando la persona que tiene el complemento cambia F. Las resoluciones y motivos anteriores de otros usuarios sí se usan como antecedentes. El estado de activación se guarda en el almacenamiento local del navegador, no en el libro.
- El complemento aparece en una pestaña propia llamada **Mis alertas** en la cinta superior de Excel.

## Publicar en GitHub Pages

1. Sube el contenido de esta carpeta a la raíz del repositorio `bryanalexexevair-lab/alertas-resoluciones`, o a una carpeta que vayas a publicar.
2. Activa GitHub Pages para esa ubicación. Los archivos del complemento deben quedar disponibles mediante HTTPS.
3. El manifiesto ya apunta a `https://bryanalexexevair-lab.github.io/alertas-resoluciones/`. Si publicas los archivos en una subcarpeta, actualiza todas las URL del manifiesto para incluirla.
4. Sube `manifest.xml` en Excel Online con **Inicio > Complementos > Más configuración > Cargar mi complemento**.
5. Abre el panel, activa las alertas una vez en ese libro y luego puedes cerrar el panel mientras trabajas.

GitHub Pages aloja el código estático del panel y de la ventana. El manifiesto solo apunta a esos archivos. Los datos del libro se leen dentro de Excel Online; no se mandan a GitHub ni a un servidor.

## Alcance y límites de esta versión

- La primera vez en cada libro, abre el panel desde **Mis alertas** y activa las alertas. El complemento configura un runtime compartido para que siga escuchando cambios aunque cierres el panel y vuelva a iniciarse cuando abras ese libro. Excel Online no conserva los manejadores después de cerrar o recargar el complemento; al abrir el libro de nuevo, el runtime se inicia y registra otra vez.
- La alerta aparece al confirmar un cambio en F; no impide guardar la resolución. Permite revisarla inmediatamente y cerrar el aviso con **Aceptar**.
- Las cédulas se comparan sin espacios, guiones ni signos de puntuación, y sin distinguir mayúsculas/minúsculas. Las resoluciones se reconocen tras quitar espacios y normalizar mayúsculas.
- Se espera encabezado en fila 1 y datos desde fila 2, igual que en la hoja revisada.
- El evento de cambios puede abarcar cualquier hoja, pero el complemento solo consulta su nombre. Solo lee encabezados y datos de B, E, F y G en hojas mensuales reconocidas; ignora las demás hojas.

