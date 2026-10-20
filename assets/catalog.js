const WORKBOOK_FILE = "../28082026. Total product on MiniApp.xlsx";
const PROMOTION_FILE = "../Miniapp_ Total promotion Q4.xlsx";
// Add new product images here using the model code as the key. Keep the
// filename (including its extension) so JPEG/JPG files both work.
const IMAGE_BY_MODEL = {
  AS10GDBY0: "../assets/images/AS10GDBY0.jpeg",
  AS10GDWH0: "../assets/images/AS10GDWH0.jpeg",
  AS20GPBK0: "../assets/images/AS20GPBK0.jpeg",
  AS20GPWU0: "../assets/images/AS20GPWU0.jpeg",
  AS20GSHU0: "../assets/images/AS20GSHU0.jpeg",
  AS30GGG1A: "../assets/images/AS30GGG1A.jpeg",
  AS30GGW1A: "../assets/images/AS30GGW1A.jpeg",
  AS55GGSY0: "../assets/images/AS55GGSY0.jpeg",
  AS60GHWG0: "../assets/images/AS60GHWG0.jpeg",
  AS65GDBY0: "../assets/images/AS65GDBY0.jpeg",
  AS65GDWH0: "../assets/images/AS65GDWH0.jpeg",
  FS15GPBK0: "../assets/images/FS15GPBK0.jpg",
  W60GLGG0: "../assets/images/W60GLGG0.jpeg"
};
// Official LG Vietnam product images for models whose files are not in
// assets/images. Keep these keyed by the base model so workbook updates can
// reuse them across regional suffixes.
const LG_IMAGE_BY_MODEL = {
  LBB33BLGA: "https://www.lg.com/content/dam/channel/wcms/vn/tu-lanh/tu-lanh-ngan-da-duoi/lbb33blga_abmpevn/gallery_new/NEW-350.jpg",
  LBB33BLM: "https://www.lg.com/content/dam/channel/wcms/vn/images/tu-lanh/lbb33blm_aeppevn_eavh_vn_c/gallery/S1.jpg",
  LBB33BGMAI: "https://www.lg.com/content/dam/channel/wcms/vn/rf/lbb33bgmai/gallery/bottom-freezer-vb4-gn-v389fgef-2024-gallery-thumbnail-01.jpg",
  LDT14BGA3: "https://www.lg.com/content/dam/channel/wcms/vn/images/may-say/ldt14bga3_anbpevn_eavh_vn_c/gallery/S-02.jpg",
  LDT14BLA4: "https://www.lg.com/content/dam/channel/wcms/vn/images/may-say/ldt14bla4_abmpevn_eavh_vn_c/gallery/DZ-01.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  LDT14SVA4: "https://www.lg.com/content/dam/channel/wcms/vn/dishwashers/ldt14sva4/gallery/450_b.jpg",
  LDN14BLA5: "https://www.lg.com/content/dam/channel/wcms/vn/dishwashers/ldn14bla5/gallery/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  IDC09M2: "https://www.lg.com/content/dam/channel/wcms/vn/h_a/airsolution/rac/air_conditioner/gallery/350_2.jpg",
  IDC12M2: "https://www.lg.com/content/dam/channel/wcms/vn/h_a/airsolution/rac/air_conditioner/gallery/350_2.jpg",
  IDC18M2: "https://www.lg.com/content/dam/channel/wcms/vn/h_a/airsolution/rac/air_conditioner/gallery/350_2.jpg",
  IDH09M1: "https://www.lg.com/content/dam/channel/wcms/vn/h_a/airsolution/rac/idh09m1/gallery/450_basic.jpg",
  IDH12M1: "https://www.lg.com/content/dam/channel/wcms/vn/h_a/airsolution/rac/idh12m1/2010-1.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  IDH18M1: "https://www.lg.com/content/dam/channel/wcms/vn/h_a/airsolution/rac/idh18m1/gallery/450-b.jpg",
  IDH24M1: "https://www.lg.com/content/dam/channel/wcms/vn/h_a/airsolution/rac/idh24m1/gallery/S_4.jpg",
  SC5MBR80H: "https://www.lg.com/content/dam/channel/wcms/vn/images/styler/sc5mbr80h_asbpevn_eavh_vn_c/s-2.jpg",
  WT2517NHEG: "https://www.lg.com/content/dam/channel/wcms/vn/images/may-giat/wt2517nheg_abgpevn_eavh_vn_c/gallery/new/3.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  TX2725AT9G: "https://www.lg.com/content/dam/channel/wcms/vn/images/wm/tx2725at9g/gallery/S_1.jpg",
  "86MRGB86BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/mini-rgb-evo/mrgb85/gp1/gallery/86-mrgb85/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "75MRGB86BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/mini-rgb-evo/mrgb85/gp1/gallery/75-mrgb85/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "65MRGB86BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/mini-rgb-evo/mrgb85/gp1/gallery/86-mrgb85/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  OLED83W6PSA: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/oled-evo/w6/gp1/gallery/83-w6/basic/350x350.png",
  LFD61BLGAI: "https://www.lg.com/content/dam/channel/wcms/vn/img24/tu-lanh-2024/lfd61blga/gallery1/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  LFD58BLMAI: "https://www.lg.com/content/dam/channel/wcms/vn/gv-k23ffqeb/gp1/gallery/basic/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  F61BMD: "https://www.lg.com/content/dam/channel/wcms/vn/f61bmd/gallery/DZ-01.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  F58BGD: "https://www.lg.com/content/dam/channel/wcms/vn/img24/tu-lanh-2024/f58bgd/gallery/DZ-01.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  LFB58BLMA: "https://www.lg.com/content/dam/channel/wcms/vn/rf/lfb58blma/gallery/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  F51EG: "https://www.lg.com/content/dam/channel/wcms/vn/tu-lanh/tu-lanh-french-door/f51eg/multi-door-richome-2026-gclb50fpgam-beigeglass-gallery-01-front-view-2010.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  F50BG: "https://www.lg.com/content/dam/channel/wcms/vn/tu-lanh/tu-lanh-french-door/f50bg/gallery_new/multi-door-richome-2026-gcfb48fpgam-blackglass-gallery-01-front-view-2010.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "GR-X257BG": "https://www.lg.com/content/dam/channel/wcms/vn/images/tu-lanh/gr-x257bg_aeepevn_eavh_vn_c/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  LSI63BLMA: "https://www.lg.com/content/dam/channel/wcms/vn/rf/lsi63blma/gallery/2010-1.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  T33SV: "https://www.lg.com/content/dam/channel/wcms/vn/t33sv/gallery/gallery/DZ-01.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  // The T30SV product page is temporarily unavailable in LG's image endpoint;
  // use the matching silver top-freezer T33SV product photo for this family.
  T30SV: "https://www.lg.com/content/dam/channel/wcms/vn/t33sv/gallery/gallery/DZ-01.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  T26BG: "https://www.lg.com/content/dam/channel/wcms/vn/rf/t26bg/DZ-1.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  T21BG: "https://www.lg.com/content/dam/channel/wcms/vn/rf/t21bg/gallery/2010_1.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  AS60GHWG0: "https://www.lg.com/content/dam/channel/wcms/vn/images/may-loc-khong-khi/as60ghwg0_abae_eavh_vn_c/gallery/DZ-01.jpg?w=800",
  OLED77G6PSA: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/oled-evo/g6/gp1/gallery/77-g6/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  OLED65G6PSA: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/oled-evo/g6/gp1/gallery/77-g6/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  OLED55G6PSA: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/oled-evo/g6/gp1/gallery/77-g6/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  OLED83C6PSA: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/oled-evo/c6/gp1/gallery/83-c6/gallery-add/2010x1334-vv1.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  OLED77C6PSA: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/oled-evo/c6/gp1/gallery/83-c6/gallery-add/2010x1334-vv1.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  OLED65C6PSA: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/oled-evo/c6/gp1/gallery/83-c6/gallery-add/2010x1334-vv1.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  OLED55C6PSA: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/oled-evo/c6/gp1/gallery/83-c6/gallery-add/2010x1334-vv1.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  OLED77B6PSA: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/oled/b6/gp1/gallery/77-b6/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  OLED65B6PSA: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/oled/b6/gp1/gallery/77-b6/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  OLED55B6PSA: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/oled/b6/gp1/gallery/77-b6/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "100QNED86BS": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/qned-evo/qned85/gp1/gallery/100-qned85/gallery/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "86QNED86BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/qned-evo/qned85/gp1/gallery/100-qned85/gallery/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "75QNED86BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/qned-evo/qned85/gp1/gallery/100-qned85/gallery/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "65QNED86BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/qned-evo/qned85/gp1/gallery/100-qned85/gallery/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "55QNED86BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/qned-evo/qned85/gp1/gallery/100-qned85/gallery/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "85QNED80BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/qned-evo/qned80/gp1/gallery/85-qned80/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "75QNED80BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/qned-evo/qned80/gp1/gallery/85-qned80/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "65QNED80BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/qned-evo/qned80/gp1/gallery/85-qned80/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "55QNED80BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/qned-evo/qned80/gp1/gallery/85-qned80/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "50QNED80BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/qned-evo/qned80/gp1/gallery/85-qned80/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "43QNED80BSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/qned-evo/qned80/gp1/gallery/85-qned80/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "98NU855BPSB": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/nano-4k-uhd/nu85/gp1/gallery/nu85-first/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "85NU855BPSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/nano-4k-uhd/nu85/gp1/gallery/nu85-first/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "75NU855BPSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/nano-4k-uhd/nu85/gp1/gallery/nu85-first/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "65NU855BPSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/nano-4k-uhd/nu85/gp1/gallery/nu85-first/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "55NU855BPSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/nano-4k-uhd/nu85/gp1/gallery/nu85-first/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "50NU855BPSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/nano-4k-uhd/nu85/gp1/gallery/nu85-first/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "43NU855BPSA": "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/tv/nano-4k-uhd/nu85/gp1/gallery/nu85-first/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "27LX6TDGA": "https://www.lg.com/content/dam/channel/wcms/vn/2025_ms_lg-com/tv/stanbyme-2/27lx6t/gp1/gallery/update/2010%20x%201334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  "32LX6BDGA": "https://www.lg.com/content/dam/channel/wcms/vn/2025_ms_lg-com/tv/stanbyme-2/32lx6bdga/gallery/new/g.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  FX1412S3KAV: "https://www.lg.com/content/dam/channel/wcms/id/laundry/fx1412s3ka-aebpein/gallery/2010_01.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  MS2535GIK: "https://www.lg.com/content/dam/channel/wcms/vn/images/lo-vi-song/ms2535gik_bseplvn_eavh_vn_c/gallery/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  MS3032JAS: "https://www.lg.com/content/dam/channel/wcms/vn/images/lo-vi-song/ms3032jas_bbkplvn_eavh_vn_c/gallery/S2.jpg",
  MS2032GIK: "https://www.lg.com/content/dam/channel/wcms/vn/images/lo-vi-song/ms2032gik_bseplvn_eavh_vn_c/gallery/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  MS2032GAS: "https://www.lg.com/content/dam/channel/wcms/vn/images/lo-vi-song/ms2032gas_bbkplvn_eavh_vn_c/gallery/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  S95AR: "https://www.lg.com/content/dam/channel/wcms/vn/images/sound-bar/s95ar/gallery/2010x1334.png/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  SG10TY: "https://www.lg.com/content/dam/channel/wcms/vn/images/promo_thhumb/new-7/2010/2010x1334_SG10TY.DVNMLLK.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  STAGE301: "https://www.lg.com/content/dam/channel/wcms/vn/images/promo_thhumb/new-7/2010/2010x1334_STAGE301.AVNMLBK.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800",
  STAGE501: "https://www.lg.com/content/dam/channel/wcms/1-channel/vn/ms/lgcom/2026/tv-audio-video/speaker/xboom-bluetooth-speaker/xboom-stage501/gallery/gallery/xboom-bluetooth-speaker-stage501-2026-gallery-gallery-01.jpg/jcr:content/renditions/thum-1600x1062.jpeg?w=800"
};
const PAGE = window.CATEGORY_PAGE || { key: "all", title: "Products", groups: [] };

const els = {
  search: document.getElementById("search"),
  type: document.getElementById("type"),
  sort: document.getElementById("sort"),
  plan: document.getElementById("plan"),
  reset: document.getElementById("reset"),
  grid: document.getElementById("grid"),
  empty: document.getElementById("empty"),
  status: document.getElementById("status"),
  title: document.getElementById("page-title"),
  copy: document.getElementById("page-copy")
};

let products = [];
let promotionsByModel = {};
const moneyFormat = new Intl.NumberFormat("vi-VN", { style: "currency", currency: "VND", maximumFractionDigits: 0 });
const money = value => value ? moneyFormat.format(value) : "n.a.";
const toNumber = value => Number.isFinite(+value) ? Math.round(+value) : null;
const text = (key, params) => window.I18N ? window.I18N.t(key, params) : key;

function sheetRows(workbook, name) {
  return XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: true, defval: "" });
}

function normalizeCategory(value, code) {
  if (!value || value === "#N/A") {
    return ({ ES: "Air Solution", AV: "Audio/Video", AO: "Audio/Video" })[code] || "Audio/Video";
  }
  if (value === "ES") return "Air Solution";
  if (value === "AV" || value === "AO") return "Audio/Video";
  if (value === "Appliance" || value === "Audio / Video" || value === "Other") return "Audio/Video";
  return value;
}

function routeCategory(category) {
  if (category === "Refrigerator") return "refrigerator";
  if (category === "TV") return "tv";
  if (category === "Audio/Video") return "audio-video";
  return "appliances";
}

function displayType(category, level, fallback) {
  if (category !== "Refrigerator") return level[9] || fallback;

  const family = [level[11], level[13], fallback].join(" ").toLowerCase();
  if (family.includes("french door") || family.includes("f/d")) return "FD";
  if ((level[9] || "").toLowerCase() === "bottom freezer" || fallback === "B/F") return "BF";

  return level[9] || fallback;
}

function parseWorkbook(workbook) {
  const priceRows = sheetRows(workbook, "Total MiniApp");
  const levelRows = sheetRows(workbook, "Product level");
  const levelBySuffix = new Map(levelRows.slice(3).filter(row => row[4]).map(row => [row[4], row]));
  return priceRows.slice(3).filter(row => row[3] && row[4]).map(row => {
    const level = levelBySuffix.get(row[4]) || [];
    const category = normalizeCategory(level[7], row[1]);
    return {
      productCode: row[1],
      subCode: row[2],
      model: row[3],
      suffix: row[4],
      outright: toNumber(row[5]) || toNumber(row[10]) || toNumber(row[6]) || toNumber(row[11]),
      monthly12: toNumber(row[7]),
      monthly24: toNumber(row[8]),
      monthly36: toNumber(row[9]),
      total12: toNumber(row[12]),
      total24: toNumber(row[13]),
      total36: toNumber(row[14]),
      category,
      route: routeCategory(category),
      type: displayType(category, level, row[2]),
      series: level[11] || "",
      variant: level[13] || ""
    };
  }).filter(product => PAGE.key === "all" || product.route === PAGE.key);
}

function parsePromotions(workbook) {
  const rows = sheetRows(workbook, "Detail price (2)");
  const promotions = {};
  rows.slice(3).filter(row => row[3] && row[4]).forEach(row => {
    const promotion = {
      eVoucher: toNumber(row[16]),
      gift: row[17],
      tvGift: row[18],
      airCareAccessory: row[19],
      southVoucher: toNumber(row[20]),
      mbBankVoucher: toNumber(row[21])
    };
    promotions[row[3]] = promotion;
    promotions[row[4]] = promotion;
  });
  return promotions;
}

function lgSearch(value) {
  return `https://www.lg.com/vn/search/?search=${encodeURIComponent(value)}`;
}

function imageFor(product) {
  return LG_IMAGE_BY_MODEL[product.model] || LG_IMAGE_BY_MODEL[product.suffix]
    || IMAGE_BY_MODEL[product.model] || IMAGE_BY_MODEL[product.suffix] || "";
}

function promotionFor(product) {
  return promotionsByModel[product.model] || promotionsByModel[product.suffix];
}

function promotionDetails(product) {
  const promotion = promotionFor(product);
  if (!promotion) return "";
  const details = [];
  const isTv = product.category === "TV";
  const isAirCare = /air/i.test(product.category);
  if (promotion.eVoucher > 0) details.push(`${text("eVoucher")}: ${money(promotion.eVoucher)}`);
  if (promotion.gift && promotion.gift !== "0") {
    if (!/^\d+$/.test(String(promotion.gift))) details.push(`${text("gift")}: ${promotion.gift}`);
    else if (isTv) details.push(text("giftAvailable"));
  }
  if (isTv && promotion.tvGift && promotion.tvGift !== "0") {
    const giftModel = /^\d+$/.test(String(promotion.tvGift)) ? "" : `: ${promotion.tvGift}`;
    details.push(`${text("tvGift")}${giftModel}`);
  }
  if (isAirCare && promotion.airCareAccessory && promotion.airCareAccessory !== "0") details.push(text("airCareAccessory"));
  if (promotion.southVoucher > 0) details.push(`${text("southVoucher")}: ${money(promotion.southVoucher)}`);
  if (promotion.mbBankVoucher > 0) details.push(`${text("mbBankVoucher")}: ${money(promotion.mbBankVoucher)}`);
  if (!details.length) return "";
  return `<div class="promotion-details"><strong>${text("promotion")}</strong><ul>${details.map(detail => `<li>${detail}</li>`).join("")}</ul></div>`;
}

function setOptions() {
  const currentSort = els.sort.value || "monthly36";
  const currentPlan = els.plan.value || "36";
  els.sort.innerHTML = `
    <option value="monthly36">${text("lowestMonthly")}</option>
    <option value="outright">${text("lowestOutright")}</option>
    <option value="model">${text("modelAZ")}</option>
  `;
  els.plan.innerHTML = `
    <option value="36">${text("months36")}</option>
    <option value="24">${text("months24")}</option>
    <option value="12">${text("months12")}</option>
  `;
  els.sort.value = currentSort;
  els.plan.value = currentPlan;
  renderTypes(products);
}

function renderTypes(list) {
  const current = els.type.value;
  const types = [...new Set(list.map(product => product.type).filter(Boolean))].sort();
  els.type.innerHTML = `<option value="all">${text("allTypes")}</option>` + types.map(type => `<option value="${type}">${type}</option>`).join("");
  els.type.value = types.includes(current) ? current : "all";
}

function card(product) {
  const planMonths = els.plan.value;
  const source = imageFor(product);
  const image = source
    ? `<img src="${source}" alt="${product.model}" onerror="this.remove();this.parentElement.insertAdjacentHTML('beforeend','<div class=&quot;placeholder&quot;>${text("imagePending")}<br>${product.model}</div>')">`
    : `<div class="placeholder">${text("imagePending")}<br>${product.model}</div>`;
  return `
    <article class="card">
      <div class="media"><span class="badge">${product.category}</span>${image}</div>
      <div class="body">
        <div class="kicker">${product.type || product.subCode}</div>
        <div class="model">${product.model}</div>
        <div class="meta">${product.variant || product.series || product.suffix}<br>${product.suffix}</div>
        ${promotionDetails(product)}
        <div class="price-row">
          <div class="price"><span>${text("outright")}</span><strong>${money(product.outright)}</strong></div>
          <div class="price"><span>${text("monthTotal", { months: planMonths })}</span><strong>${money(product[`total${planMonths}`])}</strong></div>
        </div>
        <div class="plans">
          <div class="plan"><span>${text("months12")}</span><strong>${money(product.monthly12)}</strong></div>
          <div class="plan"><span>${text("months24")}</span><strong>${money(product.monthly24)}</strong></div>
          <div class="plan"><span>${text("months36")}</span><strong>${money(product.monthly36)}</strong></div>
        </div>
        <div class="links">
          <a class="primary" href="${lgSearch(product.model)}" target="_blank" rel="noopener">${text("findImage")}</a>
          <a href="${lgSearch(product.suffix)}" target="_blank" rel="noopener">LG.com</a>
        </div>
      </div>
    </article>`;
}

function render() {
  const query = els.search.value.trim().toLowerCase();
  let list = products.slice();
  renderTypes(list);
  if (els.type.value !== "all") list = list.filter(product => product.type === els.type.value);
  if (query) {
    list = list.filter(product => [product.model, product.suffix, product.category, product.type, product.variant, product.series].join(" ").toLowerCase().includes(query));
  }
  const key = els.sort.value;
  list.sort((a, b) => key === "model" ? a.model.localeCompare(b.model) : (a[key] || Infinity) - (b[key] || Infinity));
  els.grid.innerHTML = list.map(card).join("");
  els.empty.hidden = list.length > 0;
  els.status.textContent = text("productsShown", { shown: list.length, total: products.length });
}

async function loadWorkbook() {
  els.title.textContent = text(PAGE.titleKey) || PAGE.title;
  els.copy.textContent = text(PAGE.copyKey) || PAGE.copy;
  setOptions();
  try {
    const response = await fetch(WORKBOOK_FILE);
    if (!response.ok) throw new Error(WORKBOOK_FILE);
    const buffer = await response.arrayBuffer();
    products = parseWorkbook(XLSX.read(buffer));
    try {
      const promotionResponse = await fetch(PROMOTION_FILE);
      if (promotionResponse.ok) {
        const promotionBuffer = await promotionResponse.arrayBuffer();
        promotionsByModel = parsePromotions(XLSX.read(promotionBuffer));
      }
    } catch (promotionError) {
      promotionsByModel = {};
    }
    setOptions();
    render();
  } catch (error) {
    els.status.textContent = text("workbookError");
    els.grid.innerHTML = `<div class="empty">${text("workbookHelp")}</div>`;
  }
}

[els.search, els.type, els.sort, els.plan].forEach(element => element.addEventListener("input", render));
els.reset.addEventListener("click", () => {
  els.search.value = "";
  els.type.value = "all";
  els.sort.value = "monthly36";
  els.plan.value = "36";
  render();
});

loadWorkbook();

document.addEventListener("languagechange", () => {
  if (els.search.dataset.i18nPlaceholder) els.search.placeholder = text(els.search.dataset.i18nPlaceholder);
  if (els.reset.dataset.i18n) els.reset.textContent = text(els.reset.dataset.i18n);
  els.title.textContent = text(PAGE.titleKey) || PAGE.title;
  els.copy.textContent = text(PAGE.copyKey) || PAGE.copy;
  setOptions();
  render();
});
