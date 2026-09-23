// Recherche instantanée d'organisations sur le site statique (index chargé à la demande).
(function () {
  const box = document.querySelector(".search");
  if (!box) return;
  const input = box.querySelector("input");
  const list = box.querySelector(".search-results");
  const appUrl = box.dataset.appUrl;
  const MAX_RESULTS = 8;
  let index = null;
  let loading = null;
  let active = -1;

  const normalize = (s) => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();

  function load() {
    if (!loading) {
      loading = fetch(box.dataset.indexUrl)
        .then((r) => r.json())
        .then((data) => {
          index = data.map((o) => ({ ...o, key: normalize([o.n, o.d, o.s].join(" ")) }));
        });
    }
    return loading;
  }

  function el(tag, attrs, text) {
    const e = document.createElement(tag);
    Object.assign(e, attrs);
    if (text !== undefined) e.textContent = text;
    return e;
  }

  function appLink(q, mode, label) {
    const li = el("li", { className: "app-link" });
    const params = new URLSearchParams({ mode, q });
    const a = el("a", { href: `${appUrl}/?${params}`, rel: "noopener" });
    a.append(el("span", {}, label), el("span", { className: "sub" }, "outil de recherche →"));
    li.append(a);
    return li;
  }

  function render() {
    const q = input.value.trim();
    list.replaceChildren();
    active = -1;
    if (q.length < 2 || !index) {
      list.hidden = true;
      return;
    }
    const terms = normalize(q).split(/\s+/);
    const hits = index
      .filter((o) => terms.every((t) => o.key.includes(t)))
      .sort((a, b) => b.a - a.a)
      .slice(0, MAX_RESULTS);
    for (const o of hits) {
      const a = el("a", { href: o.u });
      const label = o.s ? `${o.n} (${o.s})` : o.n;
      a.append(el("span", {}, label), el("span", { className: "sub" }, `${o.a.toLocaleString("fr-FR")} activités`));
      const li = el("li");
      li.append(a);
      list.append(li);
    }
    if (!hits.length) {
      const li = el("li");
      li.append(el("span", { className: "sub" }, "Aucune organisation trouvée"));
      list.append(li);
    }
    list.append(appLink(q, "objets", `Chercher « ${q} » dans le contenu des actions`));
    list.hidden = false;
  }

  function move(delta) {
    const items = [...list.querySelectorAll("li a")];
    if (!items.length) return;
    active = (active + delta + items.length) % items.length;
    items.forEach((a, i) => a.parentElement.classList.toggle("active", i === active));
    items[active].scrollIntoView({ block: "nearest" });
  }

  input.addEventListener("focus", load, { once: true });
  input.addEventListener("input", () => load().then(render));
  input.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown") { e.preventDefault(); move(1); }
    else if (e.key === "ArrowUp") { e.preventDefault(); move(-1); }
    else if (e.key === "Enter") {
      const items = [...list.querySelectorAll("li a")];
      const target = items[active >= 0 ? active : 0];
      if (target) window.location.href = target.href;
    } else if (e.key === "Escape") { list.hidden = true; }
  });
  document.addEventListener("click", (e) => { if (!box.contains(e.target)) list.hidden = true; });
})();
