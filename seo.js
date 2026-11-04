function generarMetaTags(p) {
  const metas = [
    { name: "description", content: p.Descripcion || "Producto eléctrico disponible" },
    { property: "og:title", content: p.Nombre },
    { property: "og:description", content: p.Descripcion },
    { property: "og:image", content: p.Imagen },
    { property: "og:url", content: window.location.href },
    { name: "twitter:card", content: "summary_large_image" }
  ];

  metas.forEach(metaData => {
    const meta = document.createElement("meta");
    for (let key in metaData) meta.setAttribute(key, metaData[key]);
    document.head.appendChild(meta);
  });
}