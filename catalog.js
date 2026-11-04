(function () {
  const config = window.CATALOG_CONFIG || {};
  const url = config.sharePointDownloadUrl;
  let cache = null;
  let lastLoaded = 0;

  function clean(value) { return value == null ? "" : String(value).trim(); }
  function normalizeProduct(row) {
    const p = {};
    Object.keys(row || {}).forEach(k => { p[clean(k).toLowerCase()] = row[k]; });
    return {
      Nombre: clean(p.nombre || p.producto || p.articulo),
      Marca: clean(p.marca), Modelo: clean(p.modelo || p.codigo),
      Categoria: clean(p.categoria || p.categoría), Precio: clean(p.precio),
      Imagen: clean(p.imagen || p.foto), Estado: clean(p.estado),
      Entrega: clean(p.entrega || p["tiempo de entrega"]),
      Descripción: clean(p["descripción"] || p.descripcion || p.detalle || p.descripcionproducto)
    };
  }
  async function loadCatalog(force) {
    if (!force && cache && Date.now() - lastLoaded < (Number(config.refreshMinutes) || 5) * 60000) return cache;
    if (!url) throw new Error("Falta configurar el enlace público de SharePoint en catalog-config.js.");
    let response;
    try { response = await fetch(url, { cache: "no-store" }); }
    catch (e) { throw new Error("El navegador no pudo leer el Excel desde SharePoint. Puede ser una restricción CORS de Microsoft 365. El enlace público permite ver/descargar el archivo, pero no necesariamente leerlo desde GitHub Pages. Se necesita habilitar una fuente CSV/JSON accesible o publicar los datos mediante Power Automate."); }
    if (!response.ok) throw new Error("SharePoint respondió con HTTP " + response.status + ". Verifica que el enlace permita acceso a cualquiera y que siga vigente.");
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength < 4) throw new Error("La respuesta de SharePoint está vacía o no parece ser un archivo Excel.");
    let workbook;
    try { workbook = XLSX.read(buffer, { type: "array" }); }
    catch (e) { throw new Error("No se pudo interpretar la descarga como Excel. Confirma que el enlace descargue el archivo .xlsx y no una página de inicio de sesión."); }
    const sheetName = workbook.SheetNames[Number(config.worksheetIndex) || 0];
    if (!sheetName) throw new Error("El archivo Excel no contiene hojas de datos.");
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "", raw: false });
    cache = rows.map(normalizeProduct).filter(p => p.Nombre && (!p.Estado || p.Estado.toLowerCase() === "activo"));
    lastLoaded = Date.now();
    return cache;
  }
  function slug(p) {
    return ((p.Marca || "") + "-" + (p.Modelo || p.Nombre || "")).toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "-").replace(/[^a-z0-9_-]/g, "").replace(/^-+|-+$/g, "");
  }
  function escapeHtml(value) {
    return String(value == null ? "" : value).replace(/[&<>"']/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[ch]));
  }
  function showError(container, error) {
    container.innerHTML = `<div class="col-12"><div class="alert alert-warning"><strong>No se pudo cargar el catálogo.</strong><p class="mb-2 mt-2">${escapeHtml(error.message || error)}</p><a class="btn btn-outline-primary btn-sm" href="${escapeHtml(url || "#")}" target="_blank" rel="noopener">Abrir catálogo de SharePoint</a></div></div>`;
  }
  async function initIndex() {
    const container = document.getElementById("productos-container");
    if (!container) return;
    const cat = document.getElementById("filter-category");
    const brand = document.getElementById("filter-brand");
    const search = document.getElementById("search-box");
    container.innerHTML = '<div class="text-center py-5"><div class="spinner-border text-primary" role="status"></div><p class="mt-3">Cargando catálogo desde SharePoint…</p></div>';
    try {
      const products = await loadCatalog();
      [...new Set(products.map(p => p.Categoria).filter(Boolean))].sort().forEach(v => cat.add(new Option(v, v)));
      [...new Set(products.map(p => p.Marca).filter(Boolean))].sort().forEach(v => brand.add(new Option(v, v)));
      function render() {
        const q = (search.value || "").toLowerCase().trim();
        const list = products.filter(p => (cat.value === "all" || p.Categoria === cat.value) && (brand.value === "all" || p.Marca === brand.value) && [p.Nombre,p.Marca,p.Modelo,p.Categoria].some(v => (v || "").toLowerCase().includes(q)));
        container.innerHTML = list.map((p, i) => `<div class="col-md-3 col-sm-6 fade-in" style="animation-delay:${i * .04}s"><div class="card h-100 shadow-sm"><img src="${escapeHtml(p.Imagen || 'no-image.png')}" class="card-img-top" alt="${escapeHtml(p.Nombre)}" loading="lazy" onerror="this.onerror=null;this.src='eps-logo.svg'"><div class="card-body text-center d-flex flex-column"><h5 class="card-title">${escapeHtml(p.Nombre)}</h5><p class="text-muted">${escapeHtml(p.Marca)} ${escapeHtml(p.Modelo)}</p><p class="fw-bold text-success">${p.Precio ? '$' + escapeHtml(p.Precio) : 'Consultar precio'}</p><a href="producto.html?item=${encodeURIComponent(slug(p))}" class="btn btn-outline-primary w-100 mt-auto">Ver producto</a></div></div></div>`).join("") || '<div class="col-12"><p class="text-center text-muted">No se encontraron productos.</p></div>';
      }
      cat.addEventListener("change", render); brand.addEventListener("change", render); search.addEventListener("input", render);
      render();
    } catch (e) { showError(container, e); }
  }
  async function initProduct() {
    const container = document.getElementById("product-container");
    if (!container) return;
    const itemSlug = new URLSearchParams(location.search).get("item");
    try {
      const products = await loadCatalog();
      const p = products.find(x => slug(x) === itemSlug);
      if (!p) { container.innerHTML = '<h3 class="text-center text-danger">Producto no encontrado</h3>'; return; }
      document.title = `${p.Nombre} | EPS Stock`;
      document.querySelector('meta[name="description"]')?.setAttribute("content", `${p.Nombre} de ${p.Marca}. Precio ${p.Precio ? '$' + p.Precio : 'a consultar'}.`);
      document.querySelector('meta[property="og:title"]')?.setAttribute("content", p.Nombre);
      document.querySelector('meta[property="og:image"]')?.setAttribute("content", p.Imagen || "logo.png");
      container.innerHTML = `<nav aria-label="breadcrumb" class="mb-4"><ol class="breadcrumb"><li class="breadcrumb-item"><a href="index.html">Inicio</a></li><li class="breadcrumb-item active">${escapeHtml(p.Categoria || 'Producto')}</li></ol></nav><div class="row align-items-center"><div class="col-md-6 text-center"><img src="${escapeHtml(p.Imagen || 'no-image.png')}" class="img-fluid rounded shadow-sm" alt="${escapeHtml(p.Nombre)}" onerror="this.onerror=null;this.src='eps-logo.svg'"></div><div class="col-md-6"><h2 class="fw-bold">${escapeHtml(p.Nombre)}</h2><p class="text-muted">${escapeHtml(p.Marca)} ${escapeHtml(p.Modelo)}</p><p><strong>Categoría:</strong> ${escapeHtml(p.Categoria || 'N/A')}</p><p><strong>Precio de referencia:</strong> <span class="text-success fw-bold">${p.Precio ? '$' + escapeHtml(p.Precio) : 'Consultar'}</span></p><p><strong>Disponibilidad / entrega:</strong> ${escapeHtml(p.Entrega || 'Consultar disponibilidad')}</p><p>${escapeHtml(p.Descripción)}</p><div class="d-flex flex-wrap gap-2 mt-3"><a href="index.html" class="btn btn-outline-secondary"><i class="bi bi-arrow-left"></i> Volver al catálogo</a><a href="https://wa.me/593939777987" target="_blank" rel="noopener" class="btn btn-eps"><i class="bi bi-whatsapp"></i> Consultar a EPS</a></div></div></div>`;
    } catch (e) { showError(container, e); }
  }
  document.addEventListener("DOMContentLoaded", () => { initIndex(); initProduct(); });
  window.reloadProductCatalog = () => loadCatalog(true);
})();
