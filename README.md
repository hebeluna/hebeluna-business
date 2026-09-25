# Hebeluna Business Management

CRM de Hebeluna Beauty: ventas con recibo, inventario, pedidos, compras, finanzas, clientas, equipo y catálogo.

- **Página:** sitio estático en Render (este repositorio, carpeta raíz).
- **Base de datos:** el Google Sheet "Inventario Hebeluna y Rashi Studio". Las hojas Hebeluna, Rashi Studio y Movimientos son el inventario; `CRM_Datos` guarda ventas, clientas, pedidos y lo demás; `CRM_Usuarios` guarda los usuarios (contraseñas con hash).
- **API:** `Code.gs` es el Apps Script del Sheet (Extensiones → Apps Script), publicado como aplicación web. Su URL va en `config.js`.

Para actualizar la página: cambia los archivos aquí y Render la vuelve a publicar sola.
