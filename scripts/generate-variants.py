#!/usr/bin/env python3
"""Generate local homepage design variants for cenzyk.art review."""

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PREVIEW = ROOT / "preview"

BRAND = "Илай — Кузнец-Дизайнер"
TAGLINE = "Готовые решения от проекта до результата"
PHONE = "+7 952 058-92-78"
PHONE_HREF = "tel:+79520589278"
TG = "https://t.me/Cenzyk"
TG_CHANNEL = "https://t.me/CenzykT_G"
EMAIL = "ilja.samatov@yandex.ru"
CITY = "Калининград и область"

COMMON_COPY = {
    "brand": BRAND,
    "tagline": TAGLINE,
    "hero_h1": "Сначала встреча и ясность — потом металл на века",
    "hero_p": "Помогаю владельцам домов и дизайнерам воплотить желание в металле: технически верно, стилистически точно, с предсказуемым результатом. Сайт — доказательство экспертизы. Главная цель — разбор вашей задачи со мной.",
    "cta_primary": "Обсудить задачу",
    "cta_secondary": "Бесплатная 3D-примерка",
    "proof_title": "Почему мне доверяют дорогие решения",
    "services_title": "Направления",
    "process_title": "Как приходим к результату",
    "meeting_title": "Что вы получите на встрече",
    "prices_note": "Цены — ориентир «от». Точная смета после разбора задачи.",
}


def nav(prefix="../"):
    return f"""
<header class="site-header">
  <a class="brand" href="{prefix}preview/index.html">
    <span class="brand-name">Илай Саматов</span>
    <span class="brand-role">Кузнец-Дизайнер</span>
  </a>
  <nav class="nav" aria-label="Основное меню">
    <a href="{prefix}balkony/">Балконы</a>
    <a href="{prefix}zabory/">Заборы</a>
    <a href="{prefix}lestnicy/">Лестницы</a>
    <a href="{prefix}mebel/">Мебель</a>
    <a href="{prefix}3d-primerka/">3D-примерка</a>
    <a class="nav-cta" href="{prefix}kontakty/">Связаться</a>
  </nav>
</header>
"""


def footer(prefix="../"):
    return f"""
<footer class="site-footer">
  <div class="footer-grid">
    <div>
      <p class="brand-name">Илай Саматов</p>
      <p class="muted">Кузнец-Дизайнер · металл под ключ</p>
      <p class="muted">{COMMON_COPY['tagline']}</p>
    </div>
    <div>
      <p><a href="{PHONE_HREF}">{PHONE}</a></p>
      <p><a href="mailto:{EMAIL}">{EMAIL}</a></p>
      <p class="muted">{CITY}</p>
    </div>
    <div>
      <p><a href="{TG}" rel="noopener">Telegram</a></p>
      <p><a href="{TG_CHANNEL}" rel="noopener">Канал с работами</a></p>
      <p><a href="{prefix}privacy/">Политика ПДн</a></p>
    </div>
  </div>
  <p class="copy">© Саматов И. В., 2026</p>
</footer>
"""


SECTIONS_BODY = """
<section class="section services" id="services">
  <div class="wrap">
    <p class="eyebrow">Направления</p>
    <h2>Заборы, балконы, лестницы, мебель</h2>
    <p class="lead">Сезонный пик — ограждения. Мебель и интерьер держат мастерскую занятой круглый год.</p>
    <div class="service-grid">
      <a class="service" href="../balkony/">
        <h3>Балконы</h3>
        <p>Сварные и кованые ограждения. 3D-примерка до производства.</p>
        <span>от 30 000 ₽/м²</span>
      </a>
      <a class="service" href="../zabory/">
        <h3>Заборы</h3>
        <p>Надёжная ограда как часть архитектуры участка — под ключ.</p>
        <span>Разбор задачи</span>
      </a>
      <a class="service" href="../lestnicy/">
        <h3>Лестницы и перила</h3>
        <p>Авторские ограждения, которые становятся акцентом дома.</p>
        <span>Разбор задачи</span>
      </a>
      <a class="service" href="../mebel/">
        <h3>Мебель и интерьер</h3>
        <p>Внесезонные изделия: столы, бра, МАФ, металл в интерьере.</p>
        <span>Разбор задачи</span>
      </a>
    </div>
  </div>
</section>

<section class="section proof" id="proof">
  <div class="wrap">
    <p class="eyebrow">Экспертиза</p>
    <h2>Сайт как доказательство: почему это не «кустарщина»</h2>
    <div class="proof-grid">
      <article>
        <h3>11+ лет в металле</h3>
        <p>Кузнец, сварщик, сборщик, слесарь и рихтовщик в одном лице. Каждый проект веду лично — с командой на производстве.</p>
      </article>
      <article>
        <h3>Технолог сварочного производства</h3>
        <p>Холодная и горячая ковка, сварка, слесарная доводка. Желание клиента упаковываю в правильную технику и стиль.</p>
      </article>
      <article>
        <h3>Школа гильдии</h3>
        <p>Обучение у Эдуарда Оганисяна — президента Гильдии кузнецов и художников Калининграда, члена Гильдии кузнецов России.</p>
      </article>
      <article>
        <h3>Гарантия до 10 лет</h3>
        <p>Договор самозанятого. Оплата 60% / 30% / 10% — прозрачные этапы без кассовых сюрпризов для обеих сторон.</p>
      </article>
    </div>
  </div>
</section>

<section class="section meeting" id="meeting">
  <div class="wrap meeting-box">
    <div>
      <p class="eyebrow">Главный продукт сайта</p>
      <h2>Не «купить сразу» — а встреча, где станет ясно</h2>
      <p class="lead">Частая боль: хочется качественно, а бюджет напрягает. На разборе честно скажу, что возможно в ваших рамках, где упростить без потери вида, и стоит ли вообще запускать производство.</p>
      <ul class="checklist">
        <li>Разберём задачу, референсы и ограничения</li>
        <li>Покажем путь: эскиз → смета → изготовление → монтаж</li>
        <li>Для простых ограждений — бесплатная 3D-примерка</li>
        <li>Сложная ковка: сначала концепт и техническое решение</li>
      </ul>
    </div>
    <aside class="meeting-card">
      <h3>Форматы связи</h3>
      <a class="btn btn-primary" href="tel:+79520589278">Позвонить</a>
      <a class="btn btn-ghost" href="https://t.me/Cenzyk" rel="noopener">Telegram</a>
      <a class="btn btn-ghost" href="../3d-primerka/">Заявка на 3D</a>
      <p class="tiny">Калининград и область · под ключ</p>
    </aside>
  </div>
</section>

<section class="section process" id="process">
  <div class="wrap">
    <p class="eyebrow">Процесс</p>
    <h2>От желания до монтажа</h2>
    <ol class="steps">
      <li><strong>Разбор</strong><span>Идея, фото, размеры, стиль, бюджетный ориентир</span></li>
      <li><strong>Эскиз / 3D</strong><span>Видите результат до оплаты производства</span></li>
      <li><strong>Смета и договор</strong><span>Прозрачные цифры, гарантия до 10 лет</span></li>
      <li><strong>Изготовление</strong><span>60% старт · 30% до покраски · 10% после монтажа</span></li>
    </ol>
  </div>
</section>

<section class="section cta-final">
  <div class="wrap cta-final-inner">
    <h2>Есть задача по металлу?</h2>
    <p>Напишите или позвоните — обсудим, как сделать правильно и на века.</p>
    <div class="cta-row">
      <a class="btn btn-primary" href="../kontakty/">Обсудить задачу</a>
      <a class="btn btn-ghost" href="../3d-primerka/">3D-примерка для Директа</a>
    </div>
  </div>
</section>
"""


VARIANTS = [
    {
        "id": "v1-quiet-metal",
        "title": "V1 · Quiet Metal",
        "summary": "Светлый тихий премиум: off-white, хаки, жёлтый акцент, full-bleed hero.",
        "css": """
:root {
  --bg: #f3f1ec;
  --bg-2: #e7e4dc;
  --ink: #141613;
  --muted: #5c6158;
  --khaki: #4a5240;
  --accent: #e2b100;
  --accent-ink: #1a1600;
  --metal: #2a2e2b;
  --line: rgba(20,22,19,.14);
  --font-display: "Syne", sans-serif;
  --font-body: "IBM Plex Sans", sans-serif;
}
body { background: var(--bg); color: var(--ink); font-family: var(--font-body); }
body::before {
  content:""; position:fixed; inset:0; pointer-events:none; opacity:.35; z-index:-1;
  background-image: radial-gradient(rgba(20,22,19,.08) 0.6px, transparent 0.6px);
  background-size: 3px 3px;
}
.hero {
  min-height: 100svh; display:grid; place-items:end start;
  background:
    linear-gradient(120deg, rgba(20,22,19,.72), rgba(74,82,64,.35) 45%, rgba(20,22,19,.2)),
    url("../assets/images/hero-metal-1.jpg") center/cover no-repeat;
  color:#f7f4ec; padding: 7rem 0 4rem;
}
.hero .wrap { width:min(1120px, calc(100% - 2.5rem)); margin:0 auto; }
.hero .brand-hero { font-family:var(--font-display); font-weight:800; font-size:clamp(2.6rem, 7vw, 5.2rem); line-height:.95; letter-spacing:-.03em; max-width:12ch; }
.hero h1 { font-family:var(--font-display); font-size:clamp(1.4rem, 3vw, 2rem); font-weight:600; max-width:28ch; margin:1.2rem 0 .8rem; }
.hero p { max-width:48ch; color:rgba(247,244,236,.88); }
.accent-bar { width:72px; height:4px; background:var(--accent); margin:1.2rem 0; }
""",
        "hero_extra_class": "",
    },
    {
        "id": "v2-blueprint",
        "title": "V2 · Blueprint Editorial",
        "summary": "Редакционный чертёж: технические линии, крупная типографика, жёлтые маркеры.",
        "css": """
:root {
  --bg: #f7f5f0;
  --bg-2: #efece4;
  --ink: #111;
  --muted: #5a5a5a;
  --khaki: #3f4736;
  --accent: #f0c400;
  --accent-ink: #111;
  --metal: #222;
  --line: rgba(17,17,17,.16);
  --font-display: "Archivo", sans-serif;
  --font-body: "IBM Plex Sans", sans-serif;
}
body {
  background:
    linear-gradient(rgba(17,17,17,.035) 1px, transparent 1px),
    linear-gradient(90deg, rgba(17,17,17,.035) 1px, transparent 1px),
    var(--bg);
  background-size: 48px 48px, 48px 48px, auto;
  color: var(--ink); font-family: var(--font-body);
}
.hero {
  min-height: 100svh; display:grid; grid-template-columns: 1.1fr .9fr; gap:0;
  border-bottom:1px solid var(--line);
}
.hero-copy { padding: 8rem 1.5rem 3rem; width:min(640px, 100%); margin-left:auto; }
.hero-visual {
  background: url("../assets/images/work-3.jpg") center/cover no-repeat;
  position:relative; min-height:70vh;
}
.hero-visual::after {
  content:""; position:absolute; inset:12% 14%;
  border:1px solid rgba(240,196,0,.85); box-shadow: inset 0 0 0 1px rgba(0,0,0,.2);
}
.hero .brand-hero { font-family:var(--font-display); font-weight:800; font-size:clamp(2.4rem, 5vw, 4.4rem); line-height:.95; text-transform:uppercase; letter-spacing:-.02em; }
.hero h1 { font-size:clamp(1.25rem, 2.4vw, 1.7rem); margin:1rem 0; max-width:28ch; font-weight:600; }
.mark { display:inline-block; background:var(--accent); color:#111; padding:.1rem .35rem; font-weight:700; }
@media (max-width:900px){ .hero{grid-template-columns:1fr;} .hero-visual{min-height:46vh; order:-1;} }
""",
        "hero_extra_class": "hero-split",
    },
    {
        "id": "v3-khaki-forge",
        "title": "V3 · Khaki Forge",
        "summary": "Хаки-поле + металл: тёплый индустриальный свет, акцент на фактуре.",
        "css": """
:root {
  --bg: #ece7da;
  --bg-2: #ddd6c4;
  --ink: #1c1f18;
  --muted: #5f6458;
  --khaki: #3e4a32;
  --accent: #d9a400;
  --accent-ink: #1c1f18;
  --metal: #2b3028;
  --line: rgba(28,31,24,.14);
  --font-display: "Syne", sans-serif;
  --font-body: "Source Serif 4", serif;
}
body { background: var(--bg); color: var(--ink); font-family: var(--font-body); }
.hero {
  min-height: 100svh; display:flex; align-items:flex-end;
  background:
    linear-gradient(180deg, rgba(62,74,50,.15), rgba(28,31,24,.78)),
    url("../assets/images/work-3.jpg") center/cover no-repeat;
  color:#f6f1e4; padding: 6rem 0 3.5rem;
}
.hero .wrap { width:min(1100px, calc(100% - 2rem)); margin:0 auto; }
.hero .brand-hero { font-family:var(--font-display); font-size:clamp(2.8rem, 7vw, 5rem); font-weight:800; line-height:.92; max-width:11ch; }
.hero h1 { font-family:var(--font-display); font-size:clamp(1.35rem, 2.8vw, 1.9rem); font-weight:700; max-width:26ch; margin:1rem 0; }
.hero p { font-size:1.05rem; max-width:46ch; color:rgba(246,241,228,.9); }
.service { background: rgba(255,255,255,.42); backdrop-filter: blur(4px); }
""",
        "hero_extra_class": "",
    },
    {
        "id": "v4-signal-yellow",
        "title": "V4 · Signal Yellow",
        "summary": "Светлая база, жёлтый как сигнал действия; бренд — главный герой экрана.",
        "css": """
:root {
  --bg: #faf8f2;
  --bg-2: #f0ebe0;
  --ink: #121412;
  --muted: #61645c;
  --khaki: #505844;
  --accent: #f5c400;
  --accent-ink: #121412;
  --metal: #262926;
  --line: rgba(18,20,18,.12);
  --font-display: "Outfit", sans-serif;
  --font-body: "IBM Plex Sans", sans-serif;
}
body { background: var(--bg); color: var(--ink); font-family: var(--font-body); }
.hero {
  min-height: 100svh; display:grid; align-content:center;
  padding: 7rem 0 4rem;
  background:
    radial-gradient(circle at 80% 20%, rgba(245,196,0,.35), transparent 28%),
    linear-gradient(135deg, #faf8f2 40%, #e9e4d6 100%);
  position:relative; overflow:hidden;
}
.hero::after {
  content:""; position:absolute; right:-8%; top:12%; width:min(48vw,560px); aspect-ratio: 4/5;
  background: url("../assets/images/hero-metal-1.jpg") center/cover no-repeat;
  clip-path: polygon(12% 0, 100% 0, 100% 100%, 0 100%);
  filter: grayscale(.15) contrast(1.05);
}
.hero .wrap { width:min(1120px, calc(100% - 2rem)); margin:0 auto; position:relative; z-index:1; }
.hero .brand-hero { font-family:var(--font-display); font-weight:800; font-size:clamp(3rem, 8vw, 5.6rem); line-height:.9; letter-spacing:-.04em; max-width:10ch; }
.hero h1 { font-size:clamp(1.2rem, 2.2vw, 1.55rem); max-width:32ch; margin:1.1rem 0; font-weight:500; }
.btn-primary { box-shadow: 4px 4px 0 var(--ink); }
@media (max-width:900px){
  .hero::after { position:relative; right:auto; top:auto; width:100%; clip-path:none; margin-top:1.5rem; aspect-ratio:16/10; display:block; }
  .hero { display:block; padding-top:6rem; }
}
""",
        "hero_extra_class": "hero-signal",
    },
    {
        "id": "v5-metal-dossier",
        "title": "V5 · Metal Dossier",
        "summary": "Досье мастера: рамки, метки, портрет процесса, продажа встречи.",
        "css": """
:root {
  --bg: #f2efe8;
  --bg-2: #e5e0d4;
  --ink: #171916;
  --muted: #5b5f56;
  --khaki: #45503c;
  --accent: #e0b000;
  --accent-ink: #171916;
  --metal: #2c302c;
  --line: rgba(23,25,22,.18);
  --font-display: "Space Grotesk", sans-serif;
  --font-body: "IBM Plex Sans", sans-serif;
}
body { background: var(--bg); color: var(--ink); font-family: var(--font-body); }
.hero {
  min-height: 100svh; padding: 7rem 1rem 3rem;
  display:grid; place-items:center;
}
.dossier {
  width:min(1100px, 100%);
  border:1px solid var(--line);
  display:grid; grid-template-columns: 1.05fr .95fr;
  background: #f7f4ec;
  box-shadow: 0 24px 60px rgba(23,25,22,.08);
}
.dossier-copy { padding: clamp(1.5rem, 4vw, 3rem); border-right:1px solid var(--line); }
.dossier-visual {
  min-height: 520px;
  background:
    linear-gradient(180deg, rgba(23,25,22,.05), rgba(23,25,22,.45)),
    url("../assets/images/work-1.jpg") center/cover no-repeat;
  position:relative;
}
.dossier-meta {
  position:absolute; left:1rem; right:1rem; bottom:1rem;
  display:flex; justify-content:space-between; gap:1rem; color:#f7f4ec; font-size:.8rem; letter-spacing:.04em; text-transform:uppercase;
}
.stamp {
  display:inline-block; border:2px solid var(--accent); color:var(--khaki);
  padding:.25rem .55rem; font-family:var(--font-display); font-weight:700; letter-spacing:.06em; text-transform:uppercase; font-size:.75rem;
}
.hero .brand-hero { font-family:var(--font-display); font-size:clamp(2.2rem, 4.5vw, 3.6rem); font-weight:700; line-height:1; margin:.8rem 0; }
.hero h1 { font-size:1.25rem; font-weight:500; max-width:30ch; margin:0 0 1rem; }
@media (max-width:900px){ .dossier{grid-template-columns:1fr;} .dossier-copy{border-right:0; border-bottom:1px solid var(--line);} }
""",
        "hero_extra_class": "hero-dossier",
    },
]


BASE_CSS = """
*,*::before,*::after{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0; line-height:1.55; text-rendering:optimizeLegibility;}
a{color:inherit}
img{max-width:100%; display:block}
.wrap{width:min(1120px, calc(100% - 2rem)); margin:0 auto;}
.site-header{
  position:fixed; inset:0 0 auto 0; z-index:20;
  display:flex; justify-content:space-between; align-items:center; gap:1rem;
  padding:.9rem 1.25rem; backdrop-filter: blur(10px);
  background: color-mix(in srgb, var(--bg) 82%, transparent);
  border-bottom:1px solid var(--line);
}
.brand{text-decoration:none; display:flex; flex-direction:column; line-height:1.1}
.brand-name{font-family:var(--font-display); font-weight:800; font-size:1.05rem}
.brand-role{font-size:.72rem; color:var(--muted); letter-spacing:.04em; text-transform:uppercase}
.nav{display:flex; flex-wrap:wrap; gap:.85rem; align-items:center; font-size:.92rem}
.nav a{text-decoration:none; color:var(--muted)}
.nav a:hover{color:var(--ink)}
.nav-cta{background:var(--accent); color:var(--accent-ink)!important; padding:.45rem .8rem; font-weight:700}
.eyebrow{font-size:.75rem; letter-spacing:.12em; text-transform:uppercase; color:var(--khaki); font-weight:700; margin:0 0 .6rem}
.section{padding:5rem 0}
.section h2{font-family:var(--font-display); font-size:clamp(1.7rem, 3vw, 2.5rem); line-height:1.1; margin:0 0 1rem; letter-spacing:-.02em; max-width:22ch}
.lead{color:var(--muted); max-width:58ch; margin:0 0 2rem}
.service-grid,.proof-grid{display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:1rem}
@media (max-width:700px){ .service-grid,.proof-grid{grid-template-columns:1fr} .nav{display:none} }
.service,.proof article{
  border:1px solid var(--line); padding:1.25rem 1.2rem; text-decoration:none; background:color-mix(in srgb, var(--bg-2) 65%, white);
  transition: transform .25s ease, border-color .25s ease;
}
.service:hover,.proof article:hover{transform:translateY(-2px); border-color: color-mix(in srgb, var(--accent) 70%, var(--line));}
.service h3,.proof h3{font-family:var(--font-display); margin:0 0 .5rem; font-size:1.25rem}
.service span{display:inline-block; margin-top:1rem; font-weight:700; color:var(--khaki)}
.meeting-box{display:grid; grid-template-columns:1.3fr .7fr; gap:1.5rem; align-items:stretch}
@media (max-width:800px){.meeting-box{grid-template-columns:1fr}}
.meeting-card{border:1px solid var(--line); padding:1.25rem; background:var(--metal); color:#f4f1e8; display:flex; flex-direction:column; gap:.7rem}
.meeting-card h3{font-family:var(--font-display); margin:0}
.checklist{padding-left:1.1rem; color:var(--muted)}
.checklist li{margin:.4rem 0}
.steps{list-style:none; padding:0; margin:0; display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:1rem; counter-reset:s}
@media (max-width:800px){.steps{grid-template-columns:1fr 1fr}}
.steps li{border-top:3px solid var(--accent); padding-top:.8rem; display:flex; flex-direction:column; gap:.35rem}
.steps li strong{font-family:var(--font-display); font-size:1.1rem}
.steps li span{color:var(--muted); font-size:.95rem}
.btn{display:inline-flex; align-items:center; justify-content:center; text-decoration:none; padding:.85rem 1.15rem; font-weight:700; border:1px solid transparent; font-family:var(--font-body)}
.btn-primary{background:var(--accent); color:var(--accent-ink)}
.btn-ghost{border-color:currentColor; background:transparent}
.cta-row{display:flex; flex-wrap:wrap; gap:.75rem; margin-top:1.2rem}
.cta-final{background:var(--metal); color:#f4f1e8}
.cta-final .btn-ghost{border-color:rgba(244,241,232,.5); color:#f4f1e8}
.site-footer{padding:3rem 0 2rem; border-top:1px solid var(--line); background:var(--bg-2)}
.footer-grid{width:min(1120px, calc(100% - 2rem)); margin:0 auto; display:grid; grid-template-columns:repeat(3,1fr); gap:1.5rem}
@media (max-width:700px){.footer-grid{grid-template-columns:1fr}}
.footer-grid a{text-decoration:none}
.muted{color:var(--muted)}
.tiny{font-size:.8rem; color:rgba(244,241,232,.7)}
.copy{width:min(1120px, calc(100% - 2rem)); margin:1.5rem auto 0; color:var(--muted); font-size:.85rem}
.variant-badge{
  position:fixed; right:12px; bottom:12px; z-index:30;
  background:var(--ink); color:#fff; padding:.45rem .7rem; font-size:.75rem; letter-spacing:.04em; text-transform:uppercase;
}
.variant-badge a{color:var(--accent); margin-left:.5rem}
"""


def hero_html(v):
    if v["id"] == "v2-blueprint":
        return f"""
<section class="hero {v['hero_extra_class']}">
  <div class="hero-copy">
    <p class="eyebrow"><span class="mark">Kaliningrad</span> · metal under key</p>
    <p class="brand-hero">{BRAND}</p>
    <p class="accent-bar"></p>
    <h1>{COMMON_COPY['hero_h1']}</h1>
    <p>{COMMON_COPY['hero_p']}</p>
    <div class="cta-row">
      <a class="btn btn-primary" href="../kontakty/">{COMMON_COPY['cta_primary']}</a>
      <a class="btn btn-ghost" href="../3d-primerka/">{COMMON_COPY['cta_secondary']}</a>
    </div>
  </div>
  <div class="hero-visual" role="img" aria-label="Работа мастерской Илая Саматова"></div>
</section>
"""
    if v["id"] == "v5-metal-dossier":
        return f"""
<section class="hero {v['hero_extra_class']}">
  <div class="dossier">
    <div class="dossier-copy">
      <span class="stamp">Case open</span>
      <p class="brand-hero">{BRAND}</p>
      <p class="muted">{TAGLINE}</p>
      <h1>{COMMON_COPY['hero_h1']}</h1>
      <p>{COMMON_COPY['hero_p']}</p>
      <div class="cta-row">
        <a class="btn btn-primary" href="../kontakty/">{COMMON_COPY['cta_primary']}</a>
        <a class="btn btn-ghost" href="../3d-primerka/">{COMMON_COPY['cta_secondary']}</a>
      </div>
    </div>
    <div class="dossier-visual">
      <div class="dossier-meta"><span>Ручейная 2А</span><span>11+ лет</span><span>Гарантия 10 лет</span></div>
    </div>
  </div>
</section>
"""
    return f"""
<section class="hero {v['hero_extra_class']}">
  <div class="wrap">
    <p class="brand-hero">{BRAND}</p>
    <p class="accent-bar" aria-hidden="true"></p>
    <p class="muted" style="color:inherit;opacity:.85">{TAGLINE}</p>
    <h1>{COMMON_COPY['hero_h1']}</h1>
    <p>{COMMON_COPY['hero_p']}</p>
    <div class="cta-row">
      <a class="btn btn-primary" href="../kontakty/">{COMMON_COPY['cta_primary']}</a>
      <a class="btn btn-ghost" href="../3d-primerka/">{COMMON_COPY['cta_secondary']}</a>
    </div>
  </div>
</section>
"""


def fonts_for(v):
    # collect unique google fonts families used
    families = {
        "v1-quiet-metal": "Syne:wght@600;700;800&family=IBM+Plex+Sans:wght@400;500;600;700",
        "v2-blueprint": "Archivo:wght@500;600;700;800&family=IBM+Plex+Sans:wght@400;500;600;700",
        "v3-khaki-forge": "Syne:wght@600;700;800&family=Source+Serif+4:wght@400;600;700",
        "v4-signal-yellow": "Outfit:wght@500;600;700;800&family=IBM+Plex+Sans:wght@400;500;600;700",
        "v5-metal-dossier": "Space+Grotesk:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700",
    }
    return families[v["id"]]


def page(v):
    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <meta name="robots" content="noindex,nofollow" />
  <title>{v['title']} · {BRAND}</title>
  <meta name="description" content="{TAGLINE}. {COMMON_COPY['hero_h1']}" />
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family={fonts_for(v)}&display=swap" rel="stylesheet" />
  <style>
{BASE_CSS}
{v['css']}
  </style>
</head>
<body>
{nav('../')}
{hero_html(v)}
{SECTIONS_BODY}
{footer('../')}
<div class="variant-badge">{v['title']}<a href="./index.html">все варианты</a></div>
<script src="../assets/js/site-motion.js" defer></script>
</body>
</html>
"""


def hub():
    cards = "\n".join(
        f"""
      <a class="card" href="./{v['id']}.html">
        <strong>{v['title']}</strong>
        <span>{v['summary']}</span>
      </a>"""
        for v in VARIANTS
    )
    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <meta name="robots" content="noindex,nofollow" />
  <title>Варианты главной · {BRAND}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=IBM+Plex+Sans:wght@400;600&display=swap" rel="stylesheet" />
  <style>
    :root {{ --bg:#f3f1ec; --ink:#141613; --accent:#e2b100; --muted:#5c6158; --line:rgba(20,22,19,.14); }}
    body {{ margin:0; font-family:"IBM Plex Sans",sans-serif; background:var(--bg); color:var(--ink); }}
    main {{ width:min(980px, calc(100% - 2rem)); margin:0 auto; padding:3rem 0 4rem; }}
    h1 {{ font-family:Syne,sans-serif; font-size:clamp(2rem,5vw,3.2rem); line-height:1; margin:0 0 1rem; }}
    p {{ color:var(--muted); max-width:60ch; }}
    .grid {{ display:grid; gap:1rem; margin-top:2rem; }}
    .card {{ display:flex; flex-direction:column; gap:.45rem; padding:1.2rem 1.25rem; border:1px solid var(--line); text-decoration:none; color:inherit; background:#fffef9; transition:transform .2s ease, border-color .2s ease; }}
    .card:hover {{ transform:translateY(-2px); border-color:var(--accent); }}
    .card strong {{ font-family:Syne,sans-serif; font-size:1.25rem; }}
    .links {{ margin-top:2rem; display:flex; flex-wrap:wrap; gap:.8rem; }}
    .links a {{ color:var(--ink); font-weight:600; }}
  </style>
</head>
<body>
<main>
  <h1>Варианты дизайна главной</h1>
  <p>Локальные прототипы для выбора стилистики. Бренд: <strong>{BRAND}</strong>. Цель экрана — продать встречу и доказать экспертизу, не «дешёвый металл».</p>
  <div class="grid">{cards}
  </div>
  <div class="links">
    <a href="../3d-primerka/">Посадка 3D-примерка (Директ)</a>
    <a href="../index.html">Корень сайта</a>
    <a href="{TG_CHANNEL}">Telegram-канал</a>
  </div>
</main>
</body>
</html>
"""


def main():
    PREVIEW.mkdir(parents=True, exist_ok=True)
    (PREVIEW / "index.html").write_text(hub(), encoding="utf-8")
    for v in VARIANTS:
        (PREVIEW / f"{v['id']}.html").write_text(page(v), encoding="utf-8")
        print("wrote", v["id"])
    print("hub ok")


if __name__ == "__main__":
    main()
