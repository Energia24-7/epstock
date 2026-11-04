# EPS Stock — catálogo de equipos y suministros eléctricos

Proyecto estático para GitHub Pages con identidad visual adaptada a EPS Electric Power Solutions (https://www.eps.ec/). El catálogo conserva la carga desde el Excel compartido de SharePoint, los filtros por categoría y marca, la búsqueda y las fichas de producto.

## Publicar
1. Descomprime el ZIP.
2. Sube el contenido de `tienda2-main` al repositorio de GitHub Pages.
3. Comprueba que `catalog-config.js` conserva el enlace correcto del Excel público.
4. Abre la página publicada y valida la carga del catálogo.

## Columnas del Excel
La primera hoja debe incluir encabezados como `Nombre`, `Marca`, `Modelo`, `Categoria`, `Precio`, `Imagen`, `Estado`, `Entrega` y `Descripción`. `Nombre` es obligatorio; si se utiliza `Estado`, se mostrarán los productos cuyo estado sea `Activo` (o los que tengan el campo vacío).

## Nota sobre el acceso a SharePoint
El enlace público no garantiza que la lectura directa desde JavaScript funcione por CORS. Si el navegador bloquea la descarga, se necesita exponer los datos como CSV/JSON de solo lectura o configurar Power Automate. No se incluyen contraseñas ni credenciales privadas en el repositorio.

## Marca y contacto
La página identifica a EPS Electric Power Solutions y enlaza a su sitio oficial. Contacto comercial publicado en el sitio de EPS: ventas@eps.ec · +593 93 977 7987. El archivo `eps-logo.svg` es un distintivo tipográfico creado para esta plantilla; reemplazarlo por el logotipo oficial autorizado si la empresa dispone del archivo de marca.

La disponibilidad y los precios deben confirmarse con el equipo comercial de EPS.
