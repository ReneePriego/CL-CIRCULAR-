const FILE_NAME = "CONAPESCA_BASE_UNIFICADA_2005_2024 (2).xlsx";
const EMPRESAS_CSV_FILE = "empresas_pescados_mariscos_mexico_V2.csv";
const EMPRESAS_CSV_FALLBACK_FILES = [];
const COMPETIDORES_CSV_FILE = "competidores_clcircular_mexico.csv";
const EXPORTACIONES_FOB_CSV_FILE = "exportaciones_fob.csv";
const CRUCES_TERRESTRES_CSV_FILES = ["Cruces Terrestres.csv"];
const PUERTOS_OCEANICOS_CSV_FILES = ["Puertos Oceánicos.csv", "Puertos Oceánicos.csv", "Puertos Oceanicos.csv"];
const NO_INFO = "No disponible públicamente";
const FTL_TON_POR_CAMION = 20;
const EXPORTACIONES_FOB_FALLBACK = [
  { year: 1990, value: 467.3 },
  { year: 1991, value: 534.6 },
  { year: 1992, value: 380.9 },
  { year: 1993, value: 436.3 },
  { year: 1994, value: 410.0 },
  { year: 1995, value: 628.1 },
  { year: 1996, value: 679.5 },
  { year: 1997, value: 707.1 },
  { year: 1998, value: 642.7 },
  { year: 1999, value: 578.7 },
  { year: 2000, value: 633.1 },
  { year: 2001, value: 568.1 },
  { year: 2002, value: 504.6 },
  { year: 2003, value: 555.1 },
  { year: 2004, value: 565.1 },
  { year: 2005, value: 564.3 },
  { year: 2006, value: 589.9 },
  { year: 2007, value: 692.9 },
  { year: 2008, value: 649.1 },
  { year: 2009, value: 673.3 },
  { year: 2010, value: 645.5 },
  { year: 2011, value: 934.9 },
  { year: 2012, value: 837.6 },
  { year: 2013, value: 829.1 },
  { year: 2014, value: 922.1 },
  { year: 2015, value: 896.9 },
  { year: 2016, value: 868.3 },
  { year: 2017, value: 1006.6 },
  { year: 2018, value: 1129.6 },
  { year: 2019, value: 1108.0 },
  { year: 2020, value: 906.3 },
  { year: 2021, value: 1054.3 },
  { year: 2022, value: 1050.9 },
  { year: 2023, value: 811.2 },
  { year: 2024, value: 776.9 },
  { year: 2025, value: 749.7 },
];
const empresasFallback = [
  {
    empresa: "Grupo Pinsa",
    rol: "Exportador e importador",
    actividad:
      "Grupo integrado de pesca, procesamiento y comercialización de atún y sardina; gran flota atunera y planta procesadora.",
    productos: "Atún en conserva, atún congelado, sardina en conserva.",
    especialidad: "Atún en conserva, atún congelado, sardina en conserva.",
    ubicacion: "Mazatlán, Sinaloa / CDMX (oficina comercial).",
    aniosOperacion: "Desde ~1980 (Pinsa Comercial ~2011)",
    alcance: "México, Estados Unidos, otros países de América; exportaciones a Europa y Asia.",
    cobertura: "México, Estados Unidos, Europa y Asia",
    paginaWeb: "https://grupopinsa.mx; https://www.pinsacomercial.com.mx",
    telefono: "Tel: +52(669)5310050",
    email: "sales.pc@pinsa.com",
    linkedin: "",
    contacto: "https://grupopinsa.mx",
    contactoLink: "https://grupopinsa.mx",
    relevancia: "Alta",
    lat: 23.2494,
    lng: -106.4111,
  },
  {
    empresa: "Pesmar / Mayaland Seafood",
    rol: "Exportador e importador",
    actividad: "Productores y distribuidores de pescados y mariscos frescos y congelados del Caribe.",
    productos: "Pescados y mariscos variados, carne de cangrejo fresca y pasteurizada.",
    especialidad: "Pescados y mariscos variados, carne de cangrejo fresca y pasteurizada.",
    ubicacion: "Yucatán, México.",
    aniosOperacion: NO_INFO,
    alcance: "Mercado nacional e internacional (detalles específicos no visibles).",
    cobertura: "Mercado nacional e internacional",
    paginaWeb: "https://ipescado.com/pesmar-pescados-y-mariscos-del-caribe/",
    telefono: "Tel: 969 935 3500, 969 934 4119 y 969 934 4436",
    email: "rudy@pesmar.com.mx, ramiro.pesmar@gmail.com",
    linkedin: "",
    contacto: "Tel: 969 935 3500",
    contactoLink: "https://wa.me/529992926171",
    relevancia: "Alta",
    lat: 20.9674,
    lng: -89.5926,
  },
  {
    empresa: "Baja Shellfish Farms",
    rol: "Exportador e importador",
    actividad: "Productor sostenible de mariscos de concha en mar abierto.",
    productos: "Ostiones (Kumiai, PaiPai, El Chingón), almejas Baja Venus, mejillones.",
    especialidad: "Ostiones, almejas y mejillones.",
    ubicacion: "Costa del Pacífico, Baja California (zona Ensenada).",
    aniosOperacion: "Desde 1991.",
    alcance: "Venta a clientes nacionales e internacionales, foco en Norteamérica.",
    cobertura: "Nacional e internacional",
    paginaWeb: "https://www.bajashellfish.com",
    telefono: "Tel: (646) 178 1684",
    email: "",
    linkedin: "",
    contacto: "https://www.bajashellfish.com",
    contactoLink: "https://www.bajashellfish.com",
    relevancia: "Media-Alta",
    lat: 31.8667,
    lng: -116.5964,
  },
  {
    empresa: "Pacifico Aquaculture",
    rol: "Exportador e importador",
    actividad: "Productor líder de lobina rayada de cultivo en mar abierto.",
    productos: "Lobina rayada (ocean-raised striped bass).",
    especialidad: "Lobina rayada.",
    ubicacion: "Ensenada, Baja California, México.",
    aniosOperacion: "Años 2010s",
    alcance: "Ventas a distribuidores, retailers y restaurantes en Norteamérica.",
    cobertura: "Norteamérica",
    paginaWeb: "https://pacificoaquaculture.com",
    telefono: "Tel: 667-764-2934",
    email: "ventas@delpacificoseafoods.com",
    linkedin: "",
    contacto: "https://pacificoaquaculture.com",
    contactoLink: "https://www.pacificoaquaculture.com",
    relevancia: "Media-Alta",
    lat: 31.8715,
    lng: -116.6219,
  },
  {
    empresa: "Grupo Amatista",
    rol: "Exportador e importador",
    actividad: "Grupo mexicano productor y comercializador de pescados y mariscos para mayoreo y retail.",
    productos: "Tilapia, basa, camarón, salmón, robalo, atún, surimi, calamar, mejillones, langostinos.",
    especialidad: "Tilapia, basa, camarón, salmón, robalo y atún.",
    ubicacion: "Cancún, Quintana Roo (oficina corporativa).",
    aniosOperacion: "Desde mediados de 2000s",
    alcance: "Cobertura nacional y presencia de compras en China y Vietnam.",
    cobertura: "Nacional e internacional",
    paginaWeb: "https://grupoamatista.com",
    telefono: "",
    email: "",
    linkedin: "",
    contacto: "https://grupoamatista.com",
    contactoLink: "https://grupoamatista.com",
    relevancia: "Media",
    lat: 21.1619,
    lng: -86.8515,
  },
  {
    empresa: "Baja Marine Foods",
    rol: "Exportador e importador",
    actividad:
      "Procesador de harina y aceite de pescado y productos congelados; parte de plataforma de Baja Aqua-Farms.",
    productos: "Harina de pescado, aceite de pescado, productos de pescado congelado.",
    especialidad: "Harina y aceite de pescado, congelados.",
    ubicacion: "Ensenada (El Sauzal de Rodríguez), Baja California.",
    aniosOperacion: "Desde 2010.",
    alcance: "Exportación de ingredientes marinos y productos congelados, mercados internacionales.",
    cobertura: "Mercados internacionales",
    paginaWeb: "http://www.bajamarinefoods.com",
    telefono: "",
    email: "",
    linkedin: "",
    contacto: "http://www.bajamarinefoods.com",
    contactoLink: "http://www.bajamarinefoods.com",
    relevancia: "Media",
    lat: 31.8667,
    lng: -116.5964,
  },
  {
    empresa: "Productores del Mar de México S.A. de C.V.",
    rol: "Exportador e importador",
    actividad: "Exportadora y procesadora de pescados y mariscos con fuerte enfoque en camarón.",
    productos: "Camarón y camarón congelado.",
    especialidad: "Camarón y camarón congelado.",
    ubicacion: "Mazatlán, Sinaloa, México.",
    aniosOperacion: NO_INFO,
    alcance: "Exportación de camarón y productos del mar a mercados internacionales.",
    cobertura: "Mercados internacionales",
    paginaWeb: "http://www.shrimparadise.co",
    telefono: "Tel: +52 (669) 118 1100",
    email: "sales@shrimparadise.com",
    linkedin: "",
    contacto: "http://www.shrimparadise.co",
    contactoLink: "http://www.shrimparadise.co",
    relevancia: "Media",
    lat: 23.2494,
    lng: -106.4111,
  },
  {
    empresa: "Frutos Marinos S.A. de C.V.",
    rol: "Exportador e importador",
    actividad:
      "Comercializadora, importadora y exportadora de pescados y mariscos frescos, congelados y secos.",
    productos: "Camarón silvestre y de cultivo, calamar gigante, callo de hacha, sardina y otros mariscos.",
    especialidad: "Camarón, calamar, callo de hacha, sardina.",
    ubicacion: "Guaymas, Sonora (y razón social relacionada en Mazatlán).",
    aniosOperacion: NO_INFO,
    alcance: "Importación y exportación para distintos mercados, venta nacional.",
    cobertura: "Nacional e internacional",
    paginaWeb: "http://www.frutosmarinos.com.mx",
    telefono: "Tel: (622) 221-5555, 221-5556",
    email: "frutosmarinos@prodigy.net.mx, irmafrumar@gmail.com",
    linkedin: "",
    contacto: "http://www.frutosmarinos.com.mx",
    contactoLink: "http://www.frutosmarinos.com.mx",
    relevancia: "Media-Alta",
    lat: 27.9186,
    lng: -110.9089,
  },
  {
    empresa: "PROMAREX S.A. de C.V.",
    rol: "Exportador e importador",
    actividad:
      "Distribuidora/comercializadora de pescados y mariscos que promueve recursos pesqueros de Sinaloa.",
    productos: "Pescados y mariscos variados (enfocados al mercado saludable).",
    especialidad: "Pescados y mariscos variados.",
    ubicacion: "Mazatlán, Sinaloa, México.",
    aniosOperacion: "Desde 1996.",
    alcance: "Mercados nacional e internacional y comercio electrónico.",
    cobertura: "Nacional e internacional",
    paginaWeb: "",
    telefono: "",
    email: "",
    linkedin: "",
    contacto: NO_INFO,
    contactoLink: "",
    relevancia: "Media",
    lat: 23.2494,
    lng: -106.4111,
  },
  {
    empresa: "MAROA S.A. de C.V.",
    rol: "Exportador e importador",
    actividad:
      "Productor, procesador y distribuidor de pescados y mariscos con granjas propias y red de proveedores.",
    productos: "Camarón, tilapia y amplio portafolio de mariscos frescos y congelados.",
    especialidad: "Camarón, tilapia y mariscos frescos/congelados.",
    ubicacion: "Guadalajara, Jalisco, México.",
    aniosOperacion: "Desde 2008.",
    alcance: "Cobertura nacional (5 CEDIS) y exportación a EE. UU. vía Maroa INC en Los Ángeles.",
    cobertura: "Nacional y exportación a EE. UU.",
    paginaWeb: "https://maroa.com.mx",
    telefono: "",
    email: "info@maroa.com.mx",
    linkedin: "",
    contacto: "https://maroa.com.mx",
    contactoLink: "https://maroa.com.mx",
    relevancia: "Media",
    lat: 20.6767,
    lng: -103.3475,
  },
  {
    empresa: "Alimentos Kay - Ricamar",
    rol: "Exportador e importador",
    actividad:
      "Elaboración de alimentos de valor agregado, principalmente empanizados de pescado congelados (marca Ricamar).",
    productos: "Empanizados de pescado congelados y otros productos del mar de valor agregado.",
    especialidad: "Empanizados de pescado y valor agregado.",
    ubicacion: "Mazatlán, Sinaloa, México.",
    aniosOperacion: "Desde 1979.",
    alcance: "Cobertura nacional con presencia en cadenas de autoservicio de todo el país.",
    cobertura: "Cobertura nacional",
    paginaWeb: "https://www.alimentoskay.com.mx",
    telefono: "",
    email: "",
    linkedin: "",
    contacto: "https://www.alimentoskay.com.mx",
    contactoLink: "https://www.alimentoskay.com.mx",
    relevancia: "Media",
    lat: 23.2494,
    lng: -106.4111,
  },
];
const empresasFallbackV2 = [
  {
    empresa: "Grupo Pinsa",
    actividad: "Grupo integrado de pesca y procesamiento industrial.",
    productos: "Atún congelado (lomos) y atún en conserva.",
    certificaciones: "MSC, BRCGS, C-TPAT / OEA, FDA",
    ubicacion: "Mazatlán, Sinaloa.",
    aniosOperacion: "Desde 1980",
    alcance: "Retailers y procesadores en EE.UU.",
    rutaTerrestre: "Mazatlán -> Nogales.",
    rutaMaritima: "Mazatlán -> Long Beach / Asia / Europa.",
    cruceFronterizo: "Nogales, AZ / Laredo, TX.",
    tempRequerida: "Congelado: <= -18 C",
    volumenEstimado: "+60 a 80 FTL/mes (Volumen constante todo el año hacia EE.UU. por contratos grandes).",
    riesgoLogisticoCsv: "MODERADO: Riesgo de robo o apagado de motor.",
    ventasAnuales: "$100M - $150M+ USD",
    retencion:
      "Son el puente comercial histórico entre las granjas mexicanas y EE.UU. Agrupan la producción de decenas de cooperativas.",
    paginaWeb: "https://www.grupopinsa.mx/",
    telefono: "+52(669)5310050",
    email: "sales.pc@pinsa.com",
    relevancia: "Alta",
    lat: 23.2494,
    lng: -106.4111,
  },
  {
    empresa: "Baja Shellfish Farms",
    actividad: "Productor acuicola vertical integrado de mariscos de concha en mar abierto.",
    productos: "Ostiones, mejillones y almejas.",
    certificaciones: "HACCP (FDA), Friend of the Sea (FoS), SSL/ESG propio",
    ubicacion: "Ensenada, BC.",
    aniosOperacion: "Desde 1991",
    alcance: "California, Newport Beach.",
    rutaTerrestre: "Ensenada -> Tijuana -> San Diego, CA.",
    rutaMaritima: "Ensenada -> San Diego, CA.",
    cruceFronterizo: "Otay Mesa, Tijuana.",
    tempRequerida: "Vivo/fresco enhielado: <= 7 C | Congelado: <= -18 C",
    volumenEstimado: "+30 a 50 FTL/mes (envios frescos frecuentes hacia CA).",
    riesgoLogisticoCsv: "CRITICO: Producto vivo, ventana corta de vida util en frontera.",
    ventasAnuales: "$15M - $30M USD",
    retencion: "Alta, clientela de chefs y wholesalers premium con contratos recurrentes.",
    paginaWeb: "https://www.bajashellfish.com",
    telefono: "+52 (646) 178 1684",
    email: "sales@bajashellfish.com",
    relevancia: "Alta",
    lat: 31.8667,
    lng: -116.5964,
  },
  {
    empresa: "Baja Aqua-Farms",
    actividad: "Cultivo y engorda en mar abierto.",
    productos: "Atún aleta azul (Fresco y Súper Congelado).",
    certificaciones: "Friend of the Sea (FoS), HACCP (FDA), ESG",
    ubicacion: "Ensenada, BC.",
    aniosOperacion: "Desde 2000",
    alcance: "California (Mercado Sushi/Sashimi).",
    rutaTerrestre: "Ensenada -> Tijuana -> Los Ángeles.",
    rutaMaritima: "Ensenada -> Japón (Súper congeladores).",
    cruceFronterizo: "Otay Mesa, Tijuana.",
    tempRequerida: "Fresco: <= 4 C / Cong: <= -18 C",
    volumenEstimado: "+50 movimientos/mes (Envíos frecuentes en LTL y FTL).",
    riesgoLogisticoCsv: "CRITICO: Romper ultracongelación o pasar 4 C destruye valor comercial en horas.",
    ventasAnuales: "$80M - $120M+ USD",
    retencion:
      "Producto premium de alto valor; cadena fría estricta y cruces frecuentes hacia California.",
    paginaWeb: "https://bajaaquafarms.com/",
    telefono: "+52 (55) 61-668479",
    email: "sales@bluefina.com",
    relevancia: "Alta",
    lat: 31.8667,
    lng: -116.5964,
  },
  {
    empresa: "Grupo Acuícola Mexicano (GAM)",
    actividad: "Producción acuícola a gran escala.",
    productos: "Camarón de cultivo (blanco).",
    certificaciones: "BAP (4 Estrellas), ASC, ISO 22000, SENASICA",
    ubicacion: "Mazatlán, Sinaloa.",
    aniosOperacion: "Desde 1990",
    alcance: "Distribuidores B2B en EE.UU.",
    rutaTerrestre: "Sinaloa -> Nogales -> EE.UU.",
    rutaMaritima: "Mazatlán -> Exportación.",
    cruceFronterizo: "Nogales, AZ.",
    tempRequerida: "Congelado: <= -18 C",
    volumenEstimado: "+50 a 70 FTL/mes (con picos de temporada).",
    riesgoLogisticoCsv: "ALTO: Volumen masivo e inspecciones aleatorias constantes en frontera.",
    ventasAnuales: "$30M - $50M USD",
    retencion:
      "Escala de producción con enfoque exportador y entregas frecuentes a clientes de EE.UU.",
    paginaWeb: "https://www.grupoacuicolamexicano.com.mx/",
    telefono: "+52 (33) 312 286 98",
    email: "ventaspacifico@gbpo.com.mx",
    relevancia: "Alta",
    lat: 23.2494,
    lng: -106.4111,
  },
  {
    empresa: "Pacifico Aquaculture",
    actividad: "Productor de lobina rayada de cultivo en mar abierto.",
    productos: "Lobina rayada (Fresco).",
    certificaciones: "BAP (4 Estrellas), ASC, SQF, HACCP (FDA)",
    ubicacion: "Ensenada, BC.",
    aniosOperacion: "Desde 2010",
    alcance: "Cadenas retail (ej. Whole Foods) y restaurantes.",
    rutaTerrestre: "Ensenada -> Otay Mesa -> Los Ángeles / San Francisco.",
    rutaMaritima: "Ensenada -> Puertos Pacífico EE.UU.",
    cruceFronterizo: "Otay Mesa, Tijuana.",
    tempRequerida: "Fresco: <= 4 C",
    volumenEstimado: "+50 movimientos/mes (producto fresco con cruces casi diarios).",
    riesgoLogisticoCsv: "ALTO: Vida de anaquel corta y sensibilidad alta a tiempos de espera en frontera.",
    ventasAnuales: "$40M - $60M+ USD",
    retencion:
      "Escala institucional con contratos de exportación y operación continua hacia California.",
    paginaWeb: "https://www.pacificoaquaculture.com/",
    telefono: "+52 646 156 5088",
    email: "ventas@pacificoaquaculture.com",
    relevancia: "Alta",
    lat: 31.8667,
    lng: -116.5964,
  },
];
let empresasData = [...empresasFallbackV2];

const state = {
  rows2024: [],
  rows2023: [],
  rowsByYear: {},
  yearsAvailable: [],
  resumen: [],
  entidades2024: [],
  speciesCaptura: [],
  speciesAcuacultura: [],
  charts: {},
  empresasMap: null,
  empresasMarkers: [],
  clientesShowAll: false,
  kpiEstadosMap: null,
  kpiEstadosLayer: null,
  riesgoRouteMap: null,
  riesgoRouteLayers: [],
  riesgoMarineMap: null,
  riesgoMarineLayers: [],
  riesgoSelectedTerrestre: {},
  riesgoSelectedMaritima: {},
  cbpWaitCacheTs: 0,
  cbpWaitCacheRows: [],
  exportFobSeries: [],
  exportFobSource: "fallback",
  empresasHash: "",
  riesgosListenersBound: false,
  serieScenario: "base",
  empresasSource: "fallback",
  empresasAutoRefreshTimer: null,
  empresasRefreshInFlight: false,
  infraCruces: [],
  infraPuertos: [],
  infraSource: "fallback",
  infraSelectedNodeId: "",
  infraCrucesMap: null,
  infraCrucesMarkers: [],
  propuestaCoverageMap: null,
  competidoresMap: null,
  viabilidadScenario: "Base",
};

const VIAB_COLORS = {
  dark: "#1a6b3a",
  mid: "#2e8f4f",
  light: "#f1f8f3",
  lightB: "#cfe2d6",
  orange: "#8fcda2",
  red: "#1a6b3a",
  yellow: "#2e8f4f",
  gray: "#355264",
  grayL: "#f8fcf9",
  grayB: "#d8e7de",
  text: "#173b2a",
  textMid: "#355264",
  textDim: "#406277",
  white: "#ffffff",
};

const VIAB_RISK_COLORS = { CRITICO: "#1a6b3a", ALTO: "#2e8f4f", MODERADO: "#8fcda2" };

const VIAB_PL_BASE = [
  { yr: "2026", rev: 98515, cogs: 15869, opex: 67000, ebitda: 15646, ni: -8000, fcf: -54000, ebm: 15.9, nm: -8.1, viajes: 3069 },
  { yr: "2027", rev: 224919, cogs: 37238, opex: 55000, ebitda: 132681, ni: 52000, fcf: 108000, ebm: 59.0, nm: 23.1, viajes: 7203 },
  { yr: "2028", rev: 231667, cogs: 38355, opex: 51290, ebitda: 142022, ni: 59000, fcf: 115000, ebm: 61.3, nm: 25.5, viajes: 7419 },
  { yr: "2029", rev: 238617, cogs: 39506, opex: 52829, ebitda: 146282, ni: 61000, fcf: 119000, ebm: 61.3, nm: 25.6, viajes: 7641 },
  { yr: "2030", rev: 245776, cogs: 40691, opex: 54414, ebitda: 150671, ni: 63000, fcf: 98000, ebm: 61.3, nm: 25.6, viajes: 7870 },
  { yr: "2031", rev: 253149, cogs: 41912, opex: 56046, ebitda: 155191, ni: 65000, fcf: 126000, ebm: 61.3, nm: 25.7, viajes: 8106 },
  { yr: "2032", rev: 260743, cogs: 43169, opex: 57727, ebitda: 159847, ni: 67000, fcf: 130000, ebm: 61.3, nm: 25.7, viajes: 8349 },
  { yr: "2033", rev: 268565, cogs: 44464, opex: 59459, ebitda: 164642, ni: 69000, fcf: 134000, ebm: 61.3, nm: 25.7, viajes: 8600 },
  { yr: "2034", rev: 276622, cogs: 45798, opex: 61243, ebitda: 169581, ni: 71000, fcf: 103000, ebm: 61.3, nm: 25.7, viajes: 8858 },
  { yr: "2035", rev: 285000, cogs: 47172, opex: 63080, ebitda: 174748, ni: 73000, fcf: 141000, ebm: 61.3, nm: 25.6, viajes: 9123 },
];
const VIAB_PL_CONS = VIAB_PL_BASE.map((d) => ({
  ...d,
  rev: Math.round(d.rev * 0.8),
  ebitda: Math.round(d.ebitda * 0.55),
  ni: Math.round(d.ni * 0.6),
}));
const VIAB_PL_OPT = VIAB_PL_BASE.map((d) => ({
  ...d,
  rev: Math.round(d.rev * 1.2),
  ebitda: Math.round(d.ebitda * 1.25),
  ni: Math.round(d.ni * 1.3),
}));
const VIAB_CUM_B = [-54, -8, 107, 226, 324, 450, 580, 714, 817, 958];
const VIAB_CUM_C = [-54, -34, 30, 90, 148, 206, 264, 328, 372, 430];
const VIAB_CUM_O = [-54, 82, 230, 390, 530, 710, 900, 1100, 1250, 1500];
const VIAB_CUM_CHART = VIAB_PL_BASE.map((d, i) => ({
  yr: d.yr,
  Base: VIAB_CUM_B[i],
  Conservador: VIAB_CUM_C[i],
  Optimista: VIAB_CUM_O[i],
}));

const VIAB_CLIENTS_Y1 = [
  {
    rank: 1,
    name: "Baja Aqua-Farms",
    tier: "Ancla",
    mes: 88,
    viajes: 1060,
    precio: 30,
    rev: 31800,
    devices: 88,
    risk: "CRITICO",
    q: "Q1 2026",
    sede: "Ensenada, BC",
    cruce: "Otay Mesa",
  },
  {
    rank: 2,
    name: "Grupo Acuicola Mex.",
    tier: "Ancla",
    mes: 60,
    viajes: 720,
    precio: 30,
    rev: 21600,
    devices: 60,
    risk: "ALTO",
    q: "Q1 2026",
    sede: "Mazatlan, SIN",
    cruce: "Nogales",
  },
  {
    rank: 3,
    name: "Grupo Pinsa",
    tier: "Estrategico",
    mes: 43,
    viajes: 513,
    precio: 35,
    rev: 17955,
    devices: 43,
    risk: "MODERADO",
    q: "Q2 2026",
    sede: "Mazatlan, SIN",
    cruce: "Nogales/Laredo",
  },
  {
    rank: 4,
    name: "Baja Shellfish Farms",
    tier: "Estrategico",
    mes: 40,
    viajes: 480,
    precio: 35,
    rev: 16800,
    devices: 40,
    risk: "CRITICO",
    q: "Q2 2026",
    sede: "Ensenada, BC",
    cruce: "Otay Mesa",
  },
  {
    rank: 5,
    name: "Pacifico Aquaculture",
    tier: "Estrategico",
    mes: 25,
    viajes: 296,
    precio: 35,
    rev: 10360,
    devices: 25,
    risk: "ALTO",
    q: "Q3 2026",
    sede: "Ensenada, BC",
    cruce: "Otay Mesa",
  },
];
const VIAB_CLIENTS_Y2 = [
  { rank: 6, name: "Pesquera Asia", tier: "Ancla", mes: 107, viajes: 1284, precio: 30, rev: 38520 },
  { rank: 7, name: "Com. Mexico Americana", tier: "Ancla", mes: 81, viajes: 974, precio: 30, rev: 29220 },
  { rank: 8, name: "Import. Esp. Angelion", tier: "Ancla", mes: 57, viajes: 684, precio: 30, rev: 20520 },
  { rank: 9, name: "Quality Fish", tier: "Estrategico", mes: 50, viajes: 596, precio: 32, rev: 19072 },
  { rank: 10, name: "Punto Austral", tier: "Estrategico", mes: 50, viajes: 596, precio: 32, rev: 19072 },
];
const VIAB_INVESTMENT = [
  { label: "Personal Año 1", monto: 25000, pct: 32 },
  { label: "Legal & Fiscal (SAT)", monto: 15000, pct: 19, warn: true },
  { label: "Infraestructura Mazatlan", monto: 12000, pct: 16 },
  { label: "Capital de trabajo", monto: 10000, pct: 13 },
  { label: "Flota sensores (256x$35)", monto: 8960, pct: 12 },
  { label: "Comercial & Marketing", monto: 6000, pct: 8 },
];
const VIAB_AMORT = [
  { mes: "Ene", saldo: 31000, interes: 413, capital: 0 },
  { mes: "Feb", saldo: 68000, interes: 907, capital: 0 },
  { mes: "Mar", saldo: 76460, interes: 1027, capital: 500 },
  { mes: "Abr", saldo: 75360, interes: 1019, capital: 1100 },
  { mes: "May", saldo: 74160, interes: 1005, capital: 1200 },
  { mes: "Jun", saldo: 72860, interes: 989, capital: 1300 },
  { mes: "Jul", saldo: 71360, interes: 971, capital: 1500 },
  { mes: "Ago", saldo: 69660, interes: 951, capital: 1700 },
  { mes: "Sep", saldo: 67760, interes: 929, capital: 1900 },
  { mes: "Oct", saldo: 65660, interes: 903, capital: 2100 },
  { mes: "Nov", saldo: 63360, interes: 875, capital: 2300 },
  { mes: "Dic", saldo: 60760, interes: 845, capital: 2600 },
];
const VIAB_FISCAL = VIAB_CLIENTS_Y1.map((c) => ({
  name: c.name.split(" ")[0],
  memb: c.rev,
  art189: Math.round(c.rev * 0.3),
  art25: Math.round(c.rev * 0.21),
  recup: Math.round(c.rev * 0.51),
  neto: Math.round(c.rev * 0.49),
}));

const VIAB_FCF_STREAM = [-76960, 15646, 132681, 142022, 129636, 155191, 159847, 164642, 148546, 174748];

function viabCalcIRR(cashflows, guess = 0.1) {
  let rate = guess;
  for (let i = 0; i < 200; i += 1) {
    let npv = 0;
    let dnpv = 0;
    cashflows.forEach((cf, t) => {
      npv += cf / Math.pow(1 + rate, t);
      dnpv -= (t * cf) / Math.pow(1 + rate, t + 1);
    });
    if (!Number.isFinite(dnpv) || dnpv === 0) break;
    const next = rate - npv / dnpv;
    if (!Number.isFinite(next)) break;
    if (Math.abs(next - rate) < 1e-8) return next;
    rate = next;
  }
  return rate;
}

function viabCalcNPV(cashflows, rate) {
  return cashflows.reduce((sum, cf, t) => sum + cf / Math.pow(1 + rate, t), 0);
}

const VIAB_IRR_BASE = viabCalcIRR(VIAB_FCF_STREAM);
const VIAB_NPV_NAFIN = viabCalcNPV(VIAB_FCF_STREAM, 0.16);
const VIAB_NPV_10 = viabCalcNPV(VIAB_FCF_STREAM, 0.1);
const VIAB_OPEX_R_Y1 = ((67000 / 98515) * 100).toFixed(1);
const VIAB_OPEX_R_Y2 = ((55000 / 224919) * 100).toFixed(1);

const CBP_CACHE_TTL_MS = 5 * 60 * 1000;
const CBP_ENDPOINTS = ["https://bwt.cbp.gov/api/waittimes", "https://bwt.cbp.gov/api/bwt"];
const HERE_TRAFFIC_API_KEY = "";
const TOMTOM_TRAFFIC_API_KEY = "";
const TRAFFIC_PROVIDER_PRIORITY = ["here", "tomtom", "cbp"];
const OSRM_ROUTE_BASE_URL = "https://router.project-osrm.org/route/v1/driving";
const ROAD_DISTANCE_TIMEOUT_MS = 8000;
const RISK_API_TIMEOUT_MS = 8000;
const CBP_FETCH_TIMEOUT_MS = 3500;
const ROAD_DISTANCE_MAX_COORDS = 25;
const COLD_PROXY_DEFAULT_THRESHOLD_C = 32;
const COLD_PROXY_YEARS = 3;
const EMPRESAS_AUTO_REFRESH_MS = 10000;
const KPI_USD_MXN_RATE = 17;
const RIESGO_MESES = [
  "enero",
  "febrero",
  "marzo",
  "abril",
  "mayo",
  "junio",
  "julio",
  "agosto",
  "septiembre",
  "octubre",
  "noviembre",
  "diciembre",
];
const coldProxyArchiveCache = new Map();
const roadDistanceCache = new Map();

const cbpCrossingKeys = {
  "Nuevo Laredo (Tamaulipas) - Laredo (Texas)": ["LAREDO", "NUEVO LAREDO", "WORLD TRADE BRIDGE", "COLUMBIA"],
  "Tijuana (Baja California) - San Ysidro (California)": ["SAN YSIDRO", "OTAY MESA", "OTAY", "TIJUANA"],
  "Ciudad Juárez (Chihuahua) - El Paso (Texas)": ["EL PASO", "CIUDAD JUAREZ", "SANTA TERESA", "ZARAGOZA"],
  "Reynosa (Tamaulipas) - Pharr/McAllen (Texas)": ["PHARR", "MCALLEN", "HIDALGO", "REYNOSA", "ANZALDUAS"],
  "Matamoros (Tamaulipas) - Brownsville (Texas)": ["BROWNSVILLE", "MATAMOROS", "LOS INDIOS"],
  "Nogales (Sonora) - Nogales (Arizona)": ["NOGALES", "MARIPOSA"],
};

const palette = {
  primary: "#083b5c",
  secondary: "#00a7a0",
  tertiary: "#62d5e9",
  accent: "#ff8a42",
  softBlue: "#3f78a8",
};

const estadoBubblePalette = [
  "#046f31",
  "#0f4c81",
  "#1b998b",
  "#d97706",
  "#9d4edd",
  "#0ea5e9",
  "#ef4444",
  "#8b5e34",
  "#16a34a",
  "#f59e0b",
  "#2563eb",
  "#db2777",
];

const infraCrucesFallback = [
  {
    id: "CT-01",
    sourceType: "terrestre",
    nombre: "Nuevo Laredo (Tamaulipas) - Laredo (Texas)",
    aduanaMx: "Nuevo Laredo",
    ciudadMx: "Nuevo Laredo",
    ciudadUsa: "Laredo",
    estadoMx: "Tamaulipas",
    estadoUsa: "Texas",
    lat: 27.4956,
    lng: -99.5074,
    ftlAnual: 4200,
    ftlMensual: 350,
    tiempoFda: "6-14 hrs",
    riesgoLevel: "ALTO",
    especies: "Camarón congelado, atún lomos, sardina procesada",
    pitch: "Documentación de temperatura por cruce para evitar rechazos en frontera.",
    cTpatActivo: "Sí",
  },
  {
    id: "CT-02",
    sourceType: "terrestre",
    nombre: "Nogales (Sonora) - Nogales (Arizona)",
    aduanaMx: "Nogales",
    ciudadMx: "Nogales",
    ciudadUsa: "Nogales",
    estadoMx: "Sonora",
    estadoUsa: "Arizona",
    lat: 31.3368,
    lng: -110.9342,
    ftlAnual: 3800,
    ftlMensual: 317,
    tiempoFda: "8-16 hrs",
    riesgoLevel: "CRITICO",
    especies: "Camarón congelado IQF, atún aleta azul",
    pitch: "Certificado automático de temperatura por embarque en condiciones de calor extremo.",
    cTpatActivo: "Sí",
  },
  {
    id: "CT-03",
    sourceType: "terrestre",
    nombre: "Otay Mesa / Tijuana (BC) - San Diego (CA)",
    aduanaMx: "Otay Mesa / Tijuana",
    ciudadMx: "Tijuana",
    ciudadUsa: "San Diego",
    estadoMx: "Baja California",
    estadoUsa: "California",
    lat: 32.5532,
    lng: -116.9739,
    ftlAnual: 1500,
    ftlMensual: 125,
    tiempoFda: "4-10 hrs",
    riesgoLevel: "CRITICO",
    especies: "Atún aleta azul fresco, lobina rayada fresca",
    pitch: "Evidencia FSMA 204 automatizada para producto fresco de alta sensibilidad.",
    cTpatActivo: "Sí",
  },
  {
    id: "CT-04",
    sourceType: "terrestre",
    nombre: "Ciudad Juárez (Chihuahua) - El Paso (Texas)",
    aduanaMx: "Ciudad Juárez (Zaragoza)",
    ciudadMx: "Ciudad Juárez",
    ciudadUsa: "El Paso",
    estadoMx: "Chihuahua",
    estadoUsa: "Texas",
    lat: 31.6904,
    lng: -106.4245,
    ftlAnual: 800,
    ftlMensual: 67,
    tiempoFda: "5-12 hrs",
    riesgoLevel: "MODERADO",
    especies: "Camarón procesado y productos empacados",
    pitch: "Monitoreo de trayectos largos para detectar desviaciones antes de frontera.",
    cTpatActivo: "Sí",
  },
  {
    id: "CT-05",
    sourceType: "terrestre",
    nombre: "Reynosa (Tamaulipas) - McAllen / Pharr (Texas)",
    aduanaMx: "Reynosa / McAllen",
    ciudadMx: "Reynosa",
    ciudadUsa: "McAllen / Pharr",
    estadoMx: "Tamaulipas",
    estadoUsa: "Texas",
    lat: 26.08,
    lng: -98.2773,
    ftlAnual: 600,
    ftlMensual: 50,
    tiempoFda: "4-10 hrs",
    riesgoLevel: "MODERADO",
    especies: "Camarón del Golfo, pulpo, productos Caribe",
    pitch: "Trazabilidad completa para camarón del Golfo. Diferénciate de competidores vietnamitas con documentación de temperatura origen-destino.",
    cTpatActivo: "Sí",
  },
  {
    id: "CT-06",
    sourceType: "terrestre",
    nombre: "Mexicali (BC) - Calexico (CA)",
    aduanaMx: "Mexicali",
    ciudadMx: "Mexicali",
    ciudadUsa: "Calexico",
    estadoMx: "Baja California",
    estadoUsa: "California",
    lat: 32.6519,
    lng: -115.4777,
    ftlAnual: 300,
    ftlMensual: 25,
    tiempoFda: "4-8 hrs",
    riesgoLevel: "MODERADO",
    especies: "Atún y productos BC",
    pitch: "Misma documentación de temperatura de Otay con menor congestión operativa.",
    cTpatActivo: "Parcial",
  },
];

const infraPuertosFallback = [
  {
    id: "PO-01",
    sourceType: "oceanico",
    nombre: "Mazatlán (Sinaloa)",
    puerto: "Mazatlán",
    ciudad: "Mazatlán",
    estado: "Sinaloa",
    lat: 23.1886,
    lng: -106.4222,
    ftlAnual: null,
    tiempoFda: "No aplica (puerto oceánico)",
    riesgoLevel: "ALTO",
    especies: "Atún congelado, camarón congelado, sardina",
    pitch: "Monitoreo continuo de reefer desde salida en puerto hasta destino internacional.",
  },
  {
    id: "PO-02",
    sourceType: "oceanico",
    nombre: "Ensenada (Baja California)",
    puerto: "Ensenada",
    ciudad: "Ensenada",
    estado: "Baja California",
    lat: 31.8576,
    lng: -116.6382,
    ftlAnual: null,
    tiempoFda: "No aplica (puerto oceánico)",
    riesgoLevel: "CRITICO",
    especies: "Atún aleta azul, langosta, abulón, lobina rayada",
    pitch: "Trazabilidad ultrafría para producto sashimi de alto valor comercial.",
  },
];

const competidoresFallback = [
  {
    empresa: "SensorGO",
    tipo: "Mexicana",
    sede: "Ciudad de México (CDMX)",
    servicio: "Sensores IoT de temperatura y humedad para cadena fría en alimentos y pharma",
    propuestaValor:
      "Empresa 100% mexicana con soporte local en español e integración IoT para cadena fría.",
    modeloNegocio: "B2B con implementacion de hardware IoT y soporte operativo continuo por suscripcion.",
    sitio: "https://sensorgo.mx",
  },
  {
    empresa: "SYCOD",
    tipo: "Mexicana",
    sede: "Ciudad de México (CDMX)",
    servicio: "Soluciones IoT para monitoreo de cadena fría - temperatura humedad y conectividad 4G/5G",
    propuestaValor:
      "Integrador nacional para proyectos empresariales de cadena fría con despliegue a medida.",
    modeloNegocio: "Integrador de proyectos empresariales con ventas consultivas y despliegues a medida.",
    sitio: "https://www.sycod.com",
  },
  {
    empresa: "Sensitech (Carrier)",
    tipo: "Internacional con oficina MX",
    sede: "Edo Mex",
    servicio: "Monitoreo cadena fría y visibilidad de cadena de suministro en tiempo real",
    propuestaValor:
      "Plataforma especializada para trazabilidad de cadena fría con foco en cumplimiento regulatorio.",
    modeloNegocio: "Plataforma empresarial global con contratos corporativos y servicios de analítica.",
    sitio: "https://www.sensitech.com",
  },
  {
    empresa: "RedGPS",
    tipo: "Internacional con presencia LATAM",
    sede: "Puebla",
    servicio: "Plataforma GPS White Label con módulo de cadena fría para flotas - alertas de temperatura en tránsito",
    propuestaValor:
      "Modelo White Label B2B para que integradores ofrezcan monitoreo de temperatura en tránsito.",
    modeloNegocio: "White Label B2B para distribuidores e integradores en logística y transporte.",
    sitio: "https://www.redgps.com",
  },
];

const competidorCoordsByName = [
  { key: "SENSITECH", lat: 19.5367, lng: -99.1947 },
  { key: "SENSORGO", lat: 19.4326, lng: -99.1332 },
  { key: "SYCOD", lat: 19.4326, lng: -99.1332 },
  { key: "GEOTAB", lat: 19.4326, lng: -99.1332 },
  { key: "DIDCOM", lat: 19.4326, lng: -99.1332 },
  { key: "REDGPS", lat: 19.0414, lng: -98.2063 },
];

let competidoresData = [];

const puertoCoords = {
  "Veracruz (Veracruz)": { lat: 19.2035, lng: -96.1342 },
  "Altamira (Tamaulipas)": { lat: 22.3927, lng: -97.9395 },
  "Tampico (Tamaulipas)": { lat: 22.2177, lng: -97.8558 },
  "Progreso (Yucatán)": { lat: 21.2825, lng: -89.6636 },
  "Manzanillo (Colima)": { lat: 19.0501, lng: -104.3188 },
  "Lázaro Cárdenas (Michoacán)": { lat: 17.9571, lng: -102.1948 },
  "Ensenada (Baja California)": { lat: 31.8667, lng: -116.5964 },
  "Mazatlán (Sinaloa)": { lat: 23.2494, lng: -106.4111 },
};

const borderCrossingCoords = {
  "Nuevo Laredo (Tamaulipas) - Laredo (Texas)": { lat: 27.5036, lng: -99.5075 },
  "Tijuana (Baja California) - San Ysidro (California)": { lat: 32.5439, lng: -117.0284 },
  "Ciudad Juárez (Chihuahua) - El Paso (Texas)": { lat: 31.7397, lng: -106.485 },
  "Reynosa (Tamaulipas) - Pharr/McAllen (Texas)": { lat: 26.0922, lng: -98.277 },
  "Matamoros (Tamaulipas) - Brownsville (Texas)": { lat: 25.8773, lng: -97.5045 },
  "Nogales (Sonora) - Nogales (Arizona)": { lat: 31.3322, lng: -110.9434 },
};

const usDestinationsByCrossing = {
  "Nuevo Laredo (Tamaulipas) - Laredo (Texas)": { nombre: "Laredo, Texas", lat: 27.5306, lng: -99.4803 },
  "Tijuana (Baja California) - San Ysidro (California)": {
    nombre: "San Diego, California",
    lat: 32.7157,
    lng: -117.1611,
  },
  "Ciudad Juárez (Chihuahua) - El Paso (Texas)": { nombre: "El Paso, Texas", lat: 31.7619, lng: -106.485 },
  "Reynosa (Tamaulipas) - Pharr/McAllen (Texas)": { nombre: "McAllen, Texas", lat: 26.2034, lng: -98.23 },
  "Matamoros (Tamaulipas) - Brownsville (Texas)": { nombre: "Brownsville, Texas", lat: 25.9017, lng: -97.4975 },
  "Nogales (Sonora) - Nogales (Arizona)": { nombre: "Phoenix, Arizona", lat: 33.4484, lng: -112.074 },
};

const usMaritimeDestByPuerto = {
  "Ensenada (Baja California)": { nombre: "San Diego, California", lat: 32.7157, lng: -117.1611 },
  "Mazatlán (Sinaloa)": { nombre: "Long Beach, California", lat: 33.7701, lng: -118.1937 },
  "Manzanillo (Colima)": { nombre: "Long Beach, California", lat: 33.7701, lng: -118.1937 },
  "Lázaro Cárdenas (Michoacán)": { nombre: "Los Ángeles, California", lat: 34.0522, lng: -118.2437 },
  "Veracruz (Veracruz)": { nombre: "Houston, Texas", lat: 29.7604, lng: -95.3698 },
  "Altamira (Tamaulipas)": { nombre: "Houston, Texas", lat: 29.7604, lng: -95.3698 },
  "Tampico (Tamaulipas)": { nombre: "Houston, Texas", lat: 29.7604, lng: -95.3698 },
  "Progreso (Yucatán)": { nombre: "Miami, Florida", lat: 25.7617, lng: -80.1918 },
};

const empresaRouteCatalog = {
  "Grupo Pinsa": {
    rutaTerrestre:
      "Mazatlán -> Mex-15N -> Nogales, Son -> Nogales, AZ -> LA / Phoenix\nCDMX -> Mex-57D -> Nuevo Laredo -> Laredo, TX",
    rutaMaritima:
      "Puerto Mazatlán -> Long Beach, CA\nManzanillo -> puertos Asia / Europa",
  },
  "Pesmar / Mayaland Seafood": {
    rutaTerrestre:
      "Mérida -> Mex-180D -> Veracruz -> CDMX\nMérida -> Mex-180 -> Cancún",
    rutaMaritima:
      "Puerto Progreso -> marítimo España / Miami\nVeracruz -> marítimo Golfo",
  },
  "Baja Shellfish Farms": {
    rutaTerrestre:
      "Ensenada -> Mex-1 -> Otay Mesa / Tijuana -> San Diego, CA -> LA\nEnsenada -> Mex-1 -> Tecate -> San Diego (alterna)",
    rutaMaritima:
      "Puerto Ensenada -> San Diego, CA\nEnsenada -> Long Beach, CA",
  },
  "Pacifico Aquaculture": {
    rutaTerrestre:
      "Ensenada -> Mex-1 -> Otay Mesa / Tijuana -> San Diego -> LA / SF",
    rutaMaritima:
      "Puerto Ensenada -> San Diego, CA\nEnsenada -> puertos Pacífico EE.UU.",
  },
  "Grupo Amatista": {
    rutaTerrestre:
      "CDMX hub -> Mex-150D -> GDL / MTY\nCDMX hub -> Mex-180 -> Cancún / Mérida\nCDMX hub -> Mex-85D -> Monterrey",
    rutaMaritima:
      "Manzanillo <- importación Asia (Shanghai / Ho Chi Minh)\nLázaro Cárdenas <- importación Asia",
  },
  "Baja Marine Foods": {
    rutaTerrestre:
      "Ensenada -> Otay Mesa / Tijuana -> San Diego, CA\nEnsenada -> Mex-1 -> Tecate -> San Diego (alterna)",
    rutaMaritima:
      "Puerto Ensenada -> Asia (harina/aceite a granel)\nEnsenada -> Long Beach, CA\nEnsenada -> puertos Japon / Corea",
  },
  "Productores del Mar de México S.A. de C.V.": {
    rutaTerrestre:
      "Mazatlán -> Mex-15N -> Nogales, Son -> Nogales, AZ -> LA / Houston\nMazatlán -> Mex-15 -> CDMX (nacional)",
    rutaMaritima:
      "Puerto Mazatlán -> Long Beach, CA\nPuerto Mazatlán -> puertos Asia (camarón IQF)",
  },
  "Frutos Marinos S.A. de C.V.": {
    rutaTerrestre:
      "Guaymas -> Mex-15N -> Nogales, Son -> Nogales, AZ -> LA / Houston\nGuaymas -> Hermosillo -> Nogales (ruta corta)",
    rutaMaritima:
      "Puerto Guaymas -> Long Beach, CA\nPuerto Guaymas -> Asia (calamar gigante congelado)",
  },
  "PROMAREX S.A. de C.V.": {
    rutaTerrestre:
      "Mazatlán -> Mex-15N -> Nogales, Son -> Nogales, AZ -> LA / Phoenix\nMazatlán -> Mex-15 -> CDMX (nacional)",
    rutaMaritima:
      "Puerto Mazatlán -> Long Beach, CA\nPuerto Mazatlán -> EE.UU. Pacífico",
  },
  "MAROA S.A. de C.V.": {
    rutaTerrestre:
      "GDL -> Mex-15D -> Nogales, Son -> Nogales, AZ -> LA\nGDL -> 5 CEDIS: CDMX · MTY · GDL · CUN · TIJ",
    rutaMaritima:
      "Manzanillo <- importación Asia (tilapia, camarón)\nManzanillo -> Long Beach, CA (exportación)",
  },
  "Alimentos Kay - Ricamar": {
    rutaTerrestre:
      "Mazatlán -> Mex-15 -> GDL -> CDMX -> cadenas autoservicio\nMazatlán -> Mex-40D -> Monterrey\nMazatlán -> Mex-15 -> Hermosillo -> Mexicali -> BC",
    rutaMaritima:
      "- (distribución nacional, sin exportación marítima confirmada)",
  },
};

const routeWaypointAliases = [
  { label: "Mazatlán, Sinaloa", lat: 23.2494, lng: -106.4111, modes: ["terrestre", "maritima"], keys: ["MAZATLAN"] },
  { label: "Sinaloa", lat: 24.8091, lng: -107.394, modes: ["terrestre"], keys: ["SINALOA"] },
  { label: "Sonora", lat: 29.0729, lng: -110.9559, modes: ["terrestre"], keys: ["SONORA"] },
  { label: "CDMX", lat: 19.4326, lng: -99.1332, modes: ["terrestre"], keys: ["CDMX", "CIUDAD DE MEXICO"] },
  { label: "Nogales, Sonora", lat: 31.3012, lng: -110.9381, modes: ["terrestre"], keys: ["NOGALES SON"] },
  {
    label: "Nogales, Arizona",
    lat: 31.3404,
    lng: -110.9343,
    modes: ["terrestre"],
    keys: ["NOGALES AZ", "NOGALES ARIZONA", "NOGALES"],
  },
  { label: "Nuevo Laredo, Tamaulipas", lat: 27.4763, lng: -99.5164, modes: ["terrestre"], keys: ["NUEVO LAREDO"] },
  { label: "Laredo, Texas", lat: 27.5306, lng: -99.4803, modes: ["terrestre"], keys: ["LAREDO TX", "LAREDO TEXAS", "LAREDO"] },
  { label: "Tijuana, Baja California", lat: 32.5149, lng: -117.0382, modes: ["terrestre"], keys: ["TIJUANA"] },
  { label: "Otay Mesa, California", lat: 32.552, lng: -116.9366, modes: ["terrestre"], keys: ["OTAY MESA"] },
  { label: "Tecate, Baja California", lat: 32.5667, lng: -116.6333, modes: ["terrestre"], keys: ["TECATE"] },
  { label: "San Ysidro, California", lat: 32.5444, lng: -117.0302, modes: ["terrestre"], keys: ["SAN YSIDRO"] },
  { label: "San Diego, California", lat: 32.7157, lng: -117.1611, modes: ["terrestre", "maritima"], keys: ["SAN DIEGO"] },
  {
    label: "Los Angeles, California",
    lat: 34.0522,
    lng: -118.2437,
    modes: ["terrestre", "maritima"],
    keys: ["LOS ANGELES", "LAX"],
  },
  {
    label: "San Francisco, California",
    lat: 37.7749,
    lng: -122.4194,
    modes: ["terrestre", "maritima"],
    keys: ["SAN FRANCISCO"],
  },
  { label: "Long Beach, California", lat: 33.7701, lng: -118.1937, modes: ["maritima"], keys: ["LONG BEACH"] },
  { label: "Phoenix, Arizona", lat: 33.4484, lng: -112.074, modes: ["terrestre"], keys: ["PHOENIX"] },
  { label: "Veracruz, Veracruz", lat: 19.2035, lng: -96.1342, modes: ["terrestre", "maritima"], keys: ["VERACRUZ"] },
  { label: "Mérida, Yucatán", lat: 20.9674, lng: -89.5926, modes: ["terrestre"], keys: ["MERIDA"] },
  { label: "Cancún, Quintana Roo", lat: 21.1619, lng: -86.8515, modes: ["terrestre"], keys: ["CANCUN"] },
  { label: "Guadalajara, Jalisco", lat: 20.6767, lng: -103.3475, modes: ["terrestre"], keys: ["GDL", "GUADALAJARA"] },
  { label: "Monterrey, Nuevo León", lat: 25.6866, lng: -100.3161, modes: ["terrestre"], keys: ["MONTERREY", "MTY"] },
  { label: "Hermosillo, Sonora", lat: 29.0729, lng: -110.9559, modes: ["terrestre"], keys: ["HERMOSILLO"] },
  { label: "Guaymas, Sonora", lat: 27.9186, lng: -110.9089, modes: ["terrestre", "maritima"], keys: ["GUAYMAS"] },
  { label: "Ensenada, Baja California", lat: 31.8667, lng: -116.5964, modes: ["terrestre", "maritima"], keys: ["ENSENADA"] },
  { label: "Mexicali, Baja California", lat: 32.6245, lng: -115.4523, modes: ["terrestre"], keys: ["MEXICALI"] },
  { label: "Manzanillo, Colima", lat: 19.0501, lng: -104.3188, modes: ["terrestre", "maritima"], keys: ["MANZANILLO"] },
  { label: "Lázaro Cárdenas, Michoacán", lat: 17.9571, lng: -102.1948, modes: ["terrestre", "maritima"], keys: ["LAZARO CARDENAS"] },
  { label: "Progreso, Yucatán", lat: 21.2825, lng: -89.6636, modes: ["maritima"], keys: ["PROGRESO"] },
  { label: "Houston, Texas", lat: 29.7604, lng: -95.3698, modes: ["terrestre", "maritima"], keys: ["HOUSTON"] },
  { label: "Miami, Florida", lat: 25.7617, lng: -80.1918, modes: ["maritima"], keys: ["MIAMI"] },
  { label: "El Paso, Texas", lat: 31.7619, lng: -106.485, modes: ["terrestre"], keys: ["EL PASO"] },
  { label: "Ciudad Juárez, Chihuahua", lat: 31.6904, lng: -106.4245, modes: ["terrestre"], keys: ["JUAREZ"] },
  { label: "Reynosa, Tamaulipas", lat: 26.0922, lng: -98.277, modes: ["terrestre"], keys: ["REYNOSA"] },
  { label: "Pharr/McAllen, Texas", lat: 26.2034, lng: -98.23, modes: ["terrestre"], keys: ["PHARR", "MCALLEN"] },
  { label: "Matamoros, Tamaulipas", lat: 25.869, lng: -97.5027, modes: ["terrestre"], keys: ["MATAMOROS"] },
  { label: "Brownsville, Texas", lat: 25.9017, lng: -97.4975, modes: ["terrestre"], keys: ["BROWNSVILLE"] },
];

const estadoCoords = {
  AGUASCALIENTES: { lat: 21.8853, lng: -102.2916 },
  "BAJA CALIFORNIA": { lat: 30.8406, lng: -115.2838 },
  "BAJA CALIFORNIA SUR": { lat: 26.0444, lng: -111.6661 },
  CAMPECHE: { lat: 19.8301, lng: -90.5349 },
  COAHUILA: { lat: 27.0587, lng: -101.7068 },
  COLIMA: { lat: 19.2452, lng: -103.7241 },
  CHIAPAS: { lat: 16.7569, lng: -93.1292 },
  CHIHUAHUA: { lat: 28.6329, lng: -106.0691 },
  "CIUDAD DE MEXICO": { lat: 19.4326, lng: -99.1332 },
  DURANGO: { lat: 24.0277, lng: -104.6532 },
  GUANAJUATO: { lat: 21.019, lng: -101.2574 },
  GUERRERO: { lat: 17.4392, lng: -99.5451 },
  HIDALGO: { lat: 20.0911, lng: -98.7624 },
  JALISCO: { lat: 20.6597, lng: -103.3496 },
  MEXICO: { lat: 19.285, lng: -99.6532 },
  MICHOACAN: { lat: 19.7008, lng: -101.1844 },
  MORELOS: { lat: 18.6813, lng: -99.1013 },
  NAYARIT: { lat: 21.7514, lng: -104.8455 },
  "NUEVO LEON": { lat: 25.5922, lng: -99.9962 },
  OAXACA: { lat: 17.0732, lng: -96.7266 },
  PUEBLA: { lat: 19.0414, lng: -98.2063 },
  QUERETARO: { lat: 20.5888, lng: -100.3899 },
  "QUINTANA ROO": { lat: 18.5141, lng: -88.3038 },
  "SAN LUIS POTOSI": { lat: 22.1565, lng: -100.9855 },
  SINALOA: { lat: 24.8091, lng: -107.394 },
  SONORA: { lat: 29.0729, lng: -110.9559 },
  TABASCO: { lat: 17.8409, lng: -92.6189 },
  TAMAULIPAS: { lat: 23.7369, lng: -99.1411 },
  TLAXCALA: { lat: 19.3139, lng: -98.2404 },
  VERACRUZ: { lat: 19.1738, lng: -96.1342 },
  YUCATAN: { lat: 20.7099, lng: -89.0943 },
  ZACATECAS: { lat: 22.7709, lng: -102.5832 },
};

const estadoAlias = {
  "COAHUILA DE ZARAGOZA": "COAHUILA",
  "MICHOACAN DE OCAMPO": "MICHOACAN",
  "VERACRUZ DE IGNACIO DE LA LLAVE": "VERACRUZ",
  "QUERETARO DE ARTEAGA": "QUERETARO",
  "ESTADO DE MEXICO": "MEXICO",
  "MEXICO (EDOMEX)": "MEXICO",
  "DISTRITO FEDERAL": "CIUDAD DE MEXICO",
  CDMX: "CIUDAD DE MEXICO",
};

const transitStateHints = {
  "NUEVO LAREDO": "TAMAULIPAS",
  REYNOSA: "TAMAULIPAS",
  MATAMOROS: "TAMAULIPAS",
  TIJUANA: "BAJA CALIFORNIA",
  ENSENADA: "BAJA CALIFORNIA",
  NOGALES: "SONORA",
  MAZATLAN: "SINALOA",
  "LOS MOCHIS": "SINALOA",
  MANZANILLO: "COLIMA",
  GUADALAJARA: "JALISCO",
  CANCUN: "QUINTANA ROO",
  VERACRUZ: "VERACRUZ",
  PROGRESO: "YUCATAN",
  TAMPICO: "TAMAULIPAS",
  ALTAMIRA: "TAMAULIPAS",
  GUAYMAS: "SONORA",
  "CIUDAD JUAREZ": "CHIHUAHUA",
  JUAREZ: "CHIHUAHUA",
  CDMX: "CIUDAD DE MEXICO",
  "ESTADO DE MEXICO": "MEXICO",
  "EDO MEX": "MEXICO",
};

const sedeCoords = [
  { key: "tlalnepantla", lat: 19.5367, lng: -99.1947 },
  { key: "estado de mexico", lat: 19.5367, lng: -99.1947 },
  { key: "edo mex", lat: 19.5367, lng: -99.1947 },
  { key: "edomex", lat: 19.5367, lng: -99.1947 },
  { key: "ciudad de mexico", lat: 19.4326, lng: -99.1332 },
  { key: "mexico nacional", lat: 19.4326, lng: -99.1332 },
  { key: "mexico latam", lat: 19.4326, lng: -99.1332 },
  { key: "mazatlan", lat: 23.2494, lng: -106.4111 },
  { key: "sinaloa", lat: 24.8091, lng: -107.394 },
  { key: "yucatan", lat: 20.9674, lng: -89.5926 },
  { key: "san diego", lat: 32.7157, lng: -117.1611 },
  { key: "ensenada", lat: 31.8667, lng: -116.5964 },
  { key: "baja california", lat: 31.8667, lng: -116.5964 },
  { key: "cancun", lat: 21.1619, lng: -86.8515 },
  { key: "quintana roo", lat: 21.1619, lng: -86.8515 },
  { key: "guaymas", lat: 27.9186, lng: -110.9089 },
  { key: "sonora", lat: 29.0729, lng: -110.9559 },
  { key: "guadalajara", lat: 20.6767, lng: -103.3475 },
  { key: "jalisco", lat: 20.6767, lng: -103.3475 },
  { key: "los mochis", lat: 25.7905, lng: -108.9859 },
  { key: "cdmx", lat: 19.4326, lng: -99.1332 },
  { key: "california", lat: 32.5149, lng: -117.0382 },
  { key: "mexico", lat: 23.6345, lng: -102.5528 },
];

document.addEventListener("DOMContentLoaded", async () => {
  initTabs();
  initFooterGlossary();
  initFileUpload();
  initSerieScenarioControls();
  initClusterFeaturesToggle();
  initClusterPcaToggle();

  seedDefaultKpiData();
  try {
    renderAll();
  } catch (error) {
    console.error("Render inicial falló:", error);
    renderKpiCardsEmergencyFallback();
  }

  try {
    await loadEmpresasData();
    state.empresasHash = buildEmpresasHash(empresasData);
  } catch (error) {
    console.error("No se pudo cargar empresas en inicio:", error);
  }

  try {
    await loadCompetidoresData();
  } catch (error) {
    console.error("No se pudo cargar competidores en inicio:", error);
  }

  try {
    await loadExportacionesFobData();
  } catch (error) {
    console.error("No se pudo cargar exportaciones FOB en inicio:", error);
  }

  try {
    await loadInfraData();
  } catch (error) {
    console.error("No se pudo cargar infraestructura en inicio:", error);
  }

  try {
    renderEmpresas();
    renderPropuestaTab();
    renderViabilidadTab();
    initInfraKpi();
    initRiesgos();
  } catch (error) {
    console.error("Error inicializando vistas secundarias:", error);
  }

  try {
    state.empresasHash = "";
    await refreshEmpresasDataFromCsv();
    startEmpresasAutoRefresh();
  } catch (error) {
    console.error("Auto-refresh de empresas no disponible:", error);
  }

  try {
    renderAll();
  } catch (error) {
    console.error("Render posterior a cargas falló:", error);
    renderKpiCardsEmergencyFallback();
  }

  try {
    ensureLibs();
    const workbook = await loadWorkbookFromPath(FILE_NAME);
    loadDataAndRender(workbook);
    setStatus("Archivo base cargado correctamente.", true);
  } catch (error) {
    setStatus(
      "No se pudo abrir automáticamente el XLSX. Usa el botón 'Cargar XLSX manual'.",
      false,
    );
    console.error(error);
  }

});

function seedDefaultKpiData() {
  if (Array.isArray(state.yearsAvailable) && state.yearsAvailable.length) return;

  const baseYears = [
    { year: 2023, ton: 1_947_600.84, valorMxnMillones: 49_722.29 },
    { year: 2024, ton: 2_004_174.97, valorMxnMillones: 47_228.29 },
  ];
  const speciesMix = [
    { especie: "CAMARON", weight: 0.12 },
    { especie: "ATUN", weight: 0.2 },
    { especie: "SARDINA", weight: 0.68 },
  ];

  state.rows2024 = [];
  state.rows2023 = [];
  state.rowsByYear = {};
  state.yearsAvailable = [];
  state.resumen = [];
  state.entidades2024 = ["NACIONAL"];
  state.speciesCaptura = ["ATUN", "CAMARON", "SARDINA"];
  state.speciesAcuacultura = [];

  baseYears.forEach((entry) => {
    const rows = speciesMix.map((mix) => ({
      origen: "CAPTURA",
      especie: mix.especie,
      entidad: "NACIONAL",
      mes: "ANUAL",
      litoral: "SIN DATO",
      pesoKg: entry.ton * 1000 * mix.weight,
      valorPesos: entry.valorMxnMillones * 1_000_000 * mix.weight,
    }));

    state.rowsByYear[entry.year] = rows;
    if (entry.year === 2024) state.rows2024 = rows;
    if (entry.year === 2023) state.rows2023 = rows;
    state.resumen.push({
      ANO: entry.year,
      PESO_DESEMBARCADO_TON: entry.ton,
      VALOR_MILLONES_MXN: entry.valorMxnMillones,
    });
  });

  state.yearsAvailable = baseYears.map((item) => item.year).sort((a, b) => b - a);
}

function startEmpresasAutoRefresh() {
  if (state.empresasAutoRefreshTimer) return;

  const tick = async () => {
    if (document.hidden) return;
    if (state.empresasRefreshInFlight) return;
    state.empresasRefreshInFlight = true;
    try {
      await refreshEmpresasDataFromCsv();
    } finally {
      state.empresasRefreshInFlight = false;
    }
  };

  state.empresasAutoRefreshTimer = setInterval(tick, EMPRESAS_AUTO_REFRESH_MS);
  tick();
  window.addEventListener("beforeunload", stopEmpresasAutoRefresh);
}

function stopEmpresasAutoRefresh() {
  if (!state.empresasAutoRefreshTimer) return;
  clearInterval(state.empresasAutoRefreshTimer);
  state.empresasAutoRefreshTimer = null;
}

async function loadEmpresasData() {
  const prevData = Array.isArray(empresasData) ? [...empresasData] : [];
  try {
    const csvText = await fetchTextFromCandidates(buildEmpresasCsvCandidates());
    if (!csvText) throw new Error(`No se pudo leer ${EMPRESAS_CSV_FILE}`);
    const parsed = parseEmpresasCsv(csvText);
    if (parsed.length) {
      empresasData = parsed;
      state.empresasSource = "csv";
      return;
    }
    // Si no se pudo parsear, usa fallback V2 para mantener consistencia con index local.
    empresasData = prevData.length ? prevData : [...empresasFallbackV2];
    state.empresasSource = "fallback";
  } catch (error) {
    console.error(error);
    // Soporte para abrir index.html sin servidor: usa fallback V2.
    empresasData = prevData.length ? prevData : [...empresasFallbackV2];
    state.empresasSource = "fallback";
  }
}

function buildEmpresasCsvCandidates() {
  const stamp = Date.now();
  const fileNames = [EMPRESAS_CSV_FILE, ...EMPRESAS_CSV_FALLBACK_FILES];
  const base = [];

  fileNames.forEach((fileName) => {
    const rawName = String(fileName || "").trim();
    if (!rawName) return;
    const encodedName = encodeURI(rawName);
    base.push(rawName, `./${rawName}`, encodedName, `./${encodedName}`);

    try {
      base.push(new URL(rawName, window.location.href).href);
      base.push(new URL(encodedName, window.location.href).href);
    } catch (error) {
      // noop
    }

    try {
      if (window.location?.origin && /^https?:/i.test(window.location.origin)) {
        base.push(`${window.location.origin}/${encodedName}`);
      }
    } catch (error) {
      // noop
    }
  });

  const withBust = [];
  base.forEach((path) => {
    if (!path) return;
    withBust.push(path);
    withBust.push(`${path}${path.includes("?") ? "&" : "?"}v=${stamp}`);
  });
  return Array.from(new Set(withBust));
}

async function loadCompetidoresData() {
  const prev = Array.isArray(competidoresData) ? [...competidoresData] : [];
  try {
    const csvText = await fetchTextFromCandidates([
      COMPETIDORES_CSV_FILE,
      `./${COMPETIDORES_CSV_FILE}`,
      encodeURI(COMPETIDORES_CSV_FILE),
    ]);
    if (!csvText) throw new Error(`No se pudo leer ${COMPETIDORES_CSV_FILE}`);
    const parsed = parseCompetidoresCsv(csvText);
    competidoresData = parsed.length ? parsed : prev.length ? prev : [...competidoresFallback];
  } catch (error) {
    console.error(error);
    competidoresData = prev.length ? prev : [...competidoresFallback];
  }
}

async function loadExportacionesFobData() {
  try {
    const csvText = await fetchTextFromCandidates([
      EXPORTACIONES_FOB_CSV_FILE,
      `./${EXPORTACIONES_FOB_CSV_FILE}`,
      encodeURI(EXPORTACIONES_FOB_CSV_FILE),
    ]);
    if (!csvText) throw new Error(`No se pudo leer ${EXPORTACIONES_FOB_CSV_FILE}`);
    const parsed = parseExportacionesFobCsv(csvText);
    state.exportFobSeries = parsed.length ? parsed : [...EXPORTACIONES_FOB_FALLBACK];
    state.exportFobSource = parsed.length ? "csv" : "fallback";
  } catch (error) {
    console.error(error);
    state.exportFobSeries = [...EXPORTACIONES_FOB_FALLBACK];
    state.exportFobSource = "fallback";
  }
}

async function loadInfraData() {
  const prevCruces = Array.isArray(state.infraCruces) ? [...state.infraCruces] : [];
  const prevPuertos = Array.isArray(state.infraPuertos) ? [...state.infraPuertos] : [];
  try {
    const [crucesText, puertosText] = await Promise.all([
      fetchTextFromCandidates(buildInfraCsvCandidates(CRUCES_TERRESTRES_CSV_FILES)),
      fetchTextFromCandidates(buildInfraCsvCandidates(PUERTOS_OCEANICOS_CSV_FILES)),
    ]);

    const crucesParsed = crucesText ? parseCrucesTerrestresCsv(crucesText) : [];
    const puertosParsed = puertosText ? parsePuertosOceanicosCsv(puertosText) : [];

    state.infraCruces = crucesParsed.length ? crucesParsed : [...infraCrucesFallback];
    state.infraPuertos = puertosParsed.length ? puertosParsed : [...infraPuertosFallback];
    state.infraSource = crucesParsed.length || puertosParsed.length ? "csv" : "fallback";
  } catch (error) {
    console.error(error);
    state.infraCruces = prevCruces.length ? prevCruces : [...infraCrucesFallback];
    state.infraPuertos = prevPuertos.length ? prevPuertos : [...infraPuertosFallback];
    state.infraSource = "fallback";
  }
}

function buildInfraCsvCandidates(fileNames = []) {
  const stamp = Date.now();
  const candidates = [];
  fileNames.forEach((fileName) => {
    const raw = String(fileName || "").trim();
    if (!raw) return;
    const encoded = encodeURI(raw);
    [raw, `./${raw}`, encoded, `./${encoded}`].forEach((item) => candidates.push(item));
    try {
      candidates.push(new URL(raw, window.location.href).href);
      candidates.push(new URL(encoded, window.location.href).href);
    } catch (error) {
      // noop
    }
  });

  const withBust = [];
  Array.from(new Set(candidates))
    .filter(Boolean)
    .forEach((path) => {
      withBust.push(path);
      withBust.push(`${path}${path.includes("?") ? "&" : "?"}v=${stamp}`);
    });

  return Array.from(new Set(withBust));
}

function parseCrucesTerrestresCsv(csvText) {
  const rows = parseCsvText(csvText);
  if (!rows.length) return [];

  const headerIndex = rows.findIndex((row) => row.some((cell) => normalizeHeader(cell) === "id"));
  const validHeaderIndex = headerIndex >= 0 ? headerIndex : 0;
  const headers = rows[validHeaderIndex].map((h) => normalizeHeader(h));
  const dataRows = rows.slice(validHeaderIndex + 1);

  return dataRows
    .map((values) => {
      if (!values.some((v) => String(v || "").trim())) return null;
      const row = {};
      headers.forEach((header, idx) => {
        row[header] = cleanCell(String(values[idx] || "").trim(), { forContact: false });
      });
      return mapCruceCsvRow(row);
    })
    .filter(Boolean);
}

function parsePuertosOceanicosCsv(csvText) {
  const rows = parseCsvText(csvText);
  if (!rows.length) return [];

  const headerIndex = rows.findIndex((row) => row.some((cell) => normalizeHeader(cell) === "id"));
  const validHeaderIndex = headerIndex >= 0 ? headerIndex : 0;
  const headers = rows[validHeaderIndex].map((h) => normalizeHeader(h));
  const dataRows = rows.slice(validHeaderIndex + 1);

  return dataRows
    .map((values) => {
      if (!values.some((v) => String(v || "").trim())) return null;
      const row = {};
      headers.forEach((header, idx) => {
        row[header] = cleanCell(String(values[idx] || "").trim(), { forContact: false });
      });
      return mapPuertoCsvRow(row);
    })
    .filter(Boolean);
}

function mapCruceCsvRow(row) {
  const aduanaMx = getCsvValueLike(row, ["aduana_mx", "aduana"], "");
  const ciudadUsa = getCsvValueLike(row, ["ciudad_ee_uu", "ciudad_usa"], "");
  if (!aduanaMx && !ciudadUsa) return null;

  const estadoMx = getCsvValueLike(row, ["estado_mx", "estado"], "");
  const estadoUsa = getCsvValueLike(row, ["estado_ee_uu", "estado_usa"], "");
  const aduanaLabel = `${aduanaMx || "Aduana"}${estadoMx ? ` (${estadoMx})` : ""}`;
  const destinoLabel = `${ciudadUsa || "Destino EE.UU."}${estadoUsa ? ` (${estadoUsa})` : ""}`;
  const lat = parseFlexibleNumber(getCsvValueLike(row, ["latitud"], ""));
  const lng = parseFlexibleNumber(getCsvValueLike(row, ["longitud"], ""));

  return {
    id: getCsvValueLike(row, ["id"], `CT-${Math.random()}`),
    sourceType: "terrestre",
    nombre: `${aduanaLabel} - ${destinoLabel}`,
    aduanaMx,
    ciudadMx: getCsvValueLike(row, ["ciudad_mx"], ""),
    ciudadUsa: ciudadUsa || getCsvValueLike(row, ["ciudad_usa"], ""),
    estadoMx,
    estadoUsa,
    lat: Number.isFinite(lat) ? lat : 23.6345,
    lng: Number.isFinite(lng) ? lng : -102.5528,
    ftlAnual: parseFlexibleNumber(getCsvValueLike(row, ["ftl_mariscos_ano_est", "ftl_mariscos_año_est"], "")),
    ftlMensual: parseFlexibleNumber(getCsvValueLike(row, ["ftl_mariscos_mes_est"], "")),
    tiempoNormal: getCsvValueLike(row, ["tiempo_cruce_normal_hrs"], NO_INFO),
    tiempoAlto: getCsvValueLike(row, ["tiempo_cruce_alto_volumen_hrs"], NO_INFO),
    tiempoFda: getCsvValueLike(row, ["tiempo_cruce_con_inspeccion_fda_hrs"], NO_INFO),
    riesgoLevel: normalizeInfraRiskLevel(getCsvValueLike(row, ["riesgo_cadena_frio", "riesgo"], "MODERADO")),
    riesgoReason: getCsvValueLike(row, ["razon_del_riesgo", "razon_riesgo"], ""),
    especies: getCsvValueLike(row, ["especies_principales"], NO_INFO),
    corredoresOrigen: getCsvValueLike(row, ["corredores_origen"], ""),
    pitch: getCsvValueLike(row, ["pitch_clcircular", "pitch_cl_circular", "pitch"], NO_INFO),
    aduanaUsa: getCsvValueLike(row, ["aduana_usa_cbp", "aduana_usa"], ""),
    fastLaneDisponible: getCsvValueLike(row, ["fast_lane_disponible"], ""),
    cTpatActivo: getCsvValueLike(row, ["C-TPAT activo"], ""),
  };
}

function mapPuertoCsvRow(row) {
  const puerto = getCsvValueLike(row, ["puerto"], "");
  if (!puerto) return null;

  const ciudad = getCsvValueLike(row, ["ciudad"], "");
  const estado = getCsvValueLike(row, ["estado"], "");
  const lat = parseFlexibleNumber(getCsvValueLike(row, ["latitud"], ""));
  const lng = parseFlexibleNumber(getCsvValueLike(row, ["longitud"], ""));
  const volMariscosTon = parseFlexibleNumber(getCsvValueLike(row, ["volumen_mariscos_ton_ano_est"], ""));
  const ftlAnualCsv = parseFlexibleNumber(getCsvValueLike(row, ["ftl_mariscos_ano_est", "ftl_ano", "ftl_anual"], ""));
  const ftlAnualEst =
    Number.isFinite(ftlAnualCsv) && ftlAnualCsv > 0
      ? Math.round(ftlAnualCsv)
      : Number.isFinite(volMariscosTon) && volMariscosTon > 0
        ? Math.round(volMariscosTon / FTL_TON_POR_CAMION)
        : null;

  return {
    id: getCsvValueLike(row, ["id"], `PO-${Math.random()}`),
    sourceType: "oceanico",
    nombre: `${puerto}${estado ? ` (${estado})` : ""}`,
    puerto,
    ciudad,
    estado,
    lat: Number.isFinite(lat) ? lat : 23.6345,
    lng: Number.isFinite(lng) ? lng : -102.5528,
    ftlAnual: ftlAnualEst,
    ftlMensual: Number.isFinite(ftlAnualEst) ? Math.round(ftlAnualEst / 12) : null,
    tiempoNormal: NO_INFO,
    tiempoAlto: NO_INFO,
    tiempoFda: "No aplica (puerto oceánico)",
    riesgoLevel: normalizeInfraRiskLevel(getCsvValueLike(row, ["riesgo_cadena_frio", "riesgo"], "MODERADO")),
    riesgoReason: getCsvValueLike(row, ["razon_del_riesgo", "razon_riesgo"], ""),
    especies: getCsvValueLike(row, ["especies_principales"], NO_INFO),
    tempRequerida: getCsvValueLike(row, ["temp_requerida"], ""),
    infraestructuraFria: getCsvValueLike(row, ["infraestructura_fria"], ""),
    destinos: getCsvValueLike(row, ["principales_destinos_exportacion"], ""),
    pitch: getCsvValueLike(row, ["pitch_clcircular", "pitch_cl_circular", "pitch"], NO_INFO),
    fuente: getCsvValueLike(row, ["fuente"], ""),
  };
}

function parseFlexibleNumber(value) {
  const raw = String(value || "").trim();
  if (!raw) return NaN;
  const cleaned = raw.replace(/[^0-9.,-]/g, "");
  if (!cleaned) return NaN;
  const normalized = cleaned.includes(",") && !cleaned.includes(".")
    ? cleaned.replace(/,/g, "")
    : cleaned.replace(/,/g, "");
  const num = Number(normalized);
  return Number.isFinite(num) ? num : NaN;
}

function normalizeInfraRiskLevel(value) {
  const text = normalizeGeoKey(value);
  if (text.includes("CRITICO")) return "CRITICO";
  if (text.includes("ALTO")) return "ALTO";
  if (text.includes("MODERADO") || text.includes("MEDIO")) return "MODERADO";
  return "MODERADO";
}

function infraRiskLabel(level) {
  if (level === "CRITICO") return "CRÍTICO";
  if (level === "ALTO") return "ALTO";
  return "MODERADO";
}

function infraRiskClass(level) {
  if (level === "CRITICO") return "infra-risk-critical";
  if (level === "ALTO") return "infra-risk-high";
  return "infra-risk-moderate";
}

function infraCtpatDisplay(rawValue, sourceType = "terrestre") {
  if (sourceType !== "terrestre") {
    return {
      label: "&#10007; No disponible",
      className: "infra-ctpat-no",
    };
  }

  const normalized = normalizeGeoKey(rawValue);
  if (normalized.includes("PARCIAL") || normalized.includes("PARTIAL")) {
    return {
      label: "~ Parcial",
      className: "infra-ctpat-partial",
    };
  }
  if (normalized.includes("SI") || normalized.includes("ACTIVO") || normalized.includes("TRUE")) {
    return {
      label: "&#10003; Activo",
      className: "infra-ctpat-active",
    };
  }
  return {
    label: "&#10007; No disponible",
    className: "infra-ctpat-no",
  };
}

function parseExportacionesFobCsv(csvText) {
  const rows = readCsvRows(csvText);
  if (!rows.length) return [];

  const headerIndex = rows.findIndex((row) =>
    row.some((cell) => {
      const h = normalizeHeader(cell);
      return h === "ano" || h === "anio" || h === "year";
    }),
  );
  const validHeaderIndex = headerIndex >= 0 ? headerIndex : 0;
  const headers = rows[validHeaderIndex].map((h) => normalizeHeader(h));
  const dataRows = rows.slice(validHeaderIndex + 1);

  const byYear = new Map();
  dataRows
    .map((values) => {
      if (!values.some((v) => String(v || "").trim())) return null;
      const row = {};
      headers.forEach((header, idx) => {
        row[header] = String(values[idx] || "").trim();
      });

      const yearRaw = getCsvValueLike(row, ["ano", "anio", "year"], "", true);
      const valueRaw = getCsvValueLike(
        row,
        ["exportaciones_fob_usd_m", "exportaciones_fob_usd", "exportaciones_fob", "fob_usd_m", "valor"],
        "",
        true,
      );

      const year = Number(String(yearRaw).replace(/[^0-9.-]/g, ""));
      const value = Number(String(valueRaw).replace(/[^0-9.-]/g, ""));
      if (!Number.isFinite(year) || !Number.isFinite(value)) return null;
      return { year, value };
    })
    .filter(Boolean)
    .forEach((point) => {
      byYear.set(point.year, (byYear.get(point.year) || 0) + point.value);
    });

  const points = Array.from(byYear.entries())
    .map(([year, value]) => ({ year, value }))
    .sort((a, b) => a.year - b.year);

  return points;
}

function parseCompetidoresCsv(csvText) {
  const rows = readCsvRows(csvText);
  if (!rows.length) return [];

  const headerIndex = rows.findIndex((row) =>
    row.some((cell) => normalizeHeader(cell) === "empresa"),
  );
  const validHeaderIndex = headerIndex >= 0 ? headerIndex : 0;
  const headers = rows[validHeaderIndex].map((h) => normalizeHeader(h));
  const dataRows = rows.slice(validHeaderIndex + 1);

  return dataRows
    .map((values) => {
      if (!values.some((v) => String(v || "").trim())) return null;
      const row = {};
      headers.forEach((header, idx) => {
        row[header] = cleanCell(String(values[idx] || "").trim(), { forContact: false });
      });
      return mapCsvRowToCompetidor(row);
    })
    .filter(Boolean);
}

function mapCsvRowToCompetidor(row) {
  const empresa = getCsvValueLike(row, ["empresa"], "");
  if (!empresa) return null;
  const sede = getCsvValueLike(
    row,
    ["sede", "sede_en_mexico", "sede en mexico", "ubicacion", "estado", "presencia"],
    NO_INFO,
  );
  const ciudad = getCsvValueLike(row, ["ciudad"], NO_INFO);
  const servicio = getCsvValueLike(
    row,
    ["servicio_principal", "servicio principal", "servicio", "actividad_principal", "solucion"],
    NO_INFO,
  );
  const lat = parseFlexibleNumber(getCsvValueLike(row, ["latitud", "lat"], ""));
  const lng = parseFlexibleNumber(getCsvValueLike(row, ["longitud", "lng", "lon"], ""));
  return {
    empresa,
    tipo: getCsvValueLike(row, ["tipo", "categoria", "segmento"], NO_INFO),
    sede: sede !== NO_INFO ? sede : ciudad,
    sedeEnMexico: sede !== NO_INFO ? sede : ciudad,
    ciudad,
    servicio,
    servicioPrincipal: servicio,
    propuestaValor: getCsvValueLike(
      row,
      ["propuesta_de_valor", "propuesta_valor", "propuesta"],
      NO_INFO,
    ),
    modeloNegocio: getCsvValueLike(
      row,
      ["modelo_de_negocio", "modelo_negocio", "business_model", "modelo"],
      NO_INFO,
    ),
    lat: Number.isFinite(lat) ? lat : undefined,
    lng: Number.isFinite(lng) ? lng : undefined,
    sitio: getCsvValueLike(row, ["sitio_web", "pagina_web", "web", "website"], "", true),
    fuente: getCsvValueLike(row, ["fuente_confirmacion", "fuente"], NO_INFO),
  };
}

function parseEmpresasCsv(csvText) {
  const rows = parseCsvText(csvText);
  if (rows.length) {
    const headerIndex = rows.findIndex((row) =>
      row.some((cell) => normalizeHeader(cell) === "empresa"),
    );
    const validHeaderIndex = headerIndex >= 0 ? headerIndex : 0;
    const headers = rows[validHeaderIndex].map((h) => normalizeHeader(h));
    const dataRows = rows.slice(validHeaderIndex + 1);

    const parsed = dataRows
      .map((values) => {
        if (!values.some((v) => String(v || "").trim())) return null;
        const row = {};
        headers.forEach((header, idx) => {
          row[header] = cleanCell(String(values[idx] || "").trim(), { forContact: false });
        });
        return mapCsvRowToEmpresa(row);
      })
      .filter((item) => item && item.empresa);

    if (parsed.length) return parsed;
  }

  const fromXlsx = parseEmpresasCsvWithXlsx(csvText);
  if (fromXlsx.length) return fromXlsx;

  return [];
}

function parseEmpresasCsvWithXlsx(csvText) {
  try {
    if (typeof XLSX === "undefined") return [];
    const workbook = XLSX.read(csvText, { type: "string" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) return [];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
    if (!rawRows.length) return [];
    return rawRows
      .map((raw) => {
        const row = {};
        Object.entries(raw).forEach(([k, v]) => {
          row[normalizeHeader(k)] = cleanCell(String(v || "").trim(), { forContact: false });
        });
        return mapCsvRowToEmpresa(row);
      })
      .filter((item) => item && item.empresa);
  } catch (error) {
    console.error("Parser XLSX CSV fallback error:", error);
    return [];
  }
}

function mapCsvRowToEmpresa(row) {
  const empresa = getCsvValue(row, ["empresa"], "");
  if (!empresa) return null;
  if (normalizeHeader(empresa) === "empresa") return null;

  const ubicacion = getCsvValueLike(
    row,
    ["sede_principal", "sede_estado", "sede", "ubicacion_principal", "ubicacion"],
    "México",
  );
  const actividad = getCsvValueLike(row, ["actividad_principal", "actividad", "actividad_empresa"], NO_INFO);
  const productos = getCsvValueLike(row, ["productos_clave", "productos", "productos_principales"], NO_INFO);
  const certificaciones = getCsvValueLike(
    row,
    [
      "certificaciones",
      "certificacion",
      "certificaciones_clave",
      "certificaciones_inocuidad",
      "certificaciones_y_normativas",
      "certificaciones_normativas",
      "certificaciones_calidad",
      "certificaciones_fitosanitarias",
      "certificaciones_regulatorias",
      "normativas",
    ],
    NO_INFO,
  );
  const frecuenciaSensor = getCsvValueLike(
    row,
    [
      "frecuencia_sensor",
      "frecuencia_de_lectura",
      "frecuencia_lectura",
      "frecuencia_registro",
      "frecuencia_monitoreo",
      "intervalo_lectura",
    ],
    "",
    true,
  );
  const aniosOperacion = getCsvValueLike(
    row,
    [
      "anos_en_operacion_aprox",
      "anios_en_operacion_aprox",
      "anos_en_operacion",
      "anios_en_operacion",
      "tiempo_operacion",
    ],
    NO_INFO,
  );
  const alcance = getCsvValueLike(
    row,
    [
      "destino_mercado_usa",
      "destino_mercado",
      "mercado_usa",
      "mercado_estados_unidos",
      "exportacion_alcance",
      "exportacion_alcance_mercado",
      "exportacion",
      "alcance",
    ],
    NO_INFO,
  );
  const destinoFinal = getCsvValueLike(
    row,
    [
      "destino_final",
      "destino_final_usa",
      "destino_usa",
      "mercado_destino_usa",
      "destino_mercado_usa",
      "destino_mercado",
      "destino",
    ],
    alcance,
  );
  const paginaWeb = getCsvValueLike(row, ["pagina_web", "sitio_web", "web", "website"], "", true);
  const telefono = getCsvValueLike(row, ["telefono_whatsapp", "telefono", "whatsapp", "celular"], "", true);
  const email = getCsvValueLike(row, ["email", "e_mail", "correo", "correo_electronico"], "", true);
  const linkedin = getCsvValueLike(row, ["linkedin", "linked_in"], "", true);
  const rutaTerrestre = getCsvValueLike(row, ["ruta_terrestre", "ruta_terrestre_principal"], "", true);
  const rutaMaritima = getCsvValueLike(
    row,
    ["ruta_maritima", "ruta_maritima_principal", "ruta_oceanica"],
    "",
    true,
  );
  const cruceFronterizo = getCsvValueLike(row, ["cruce_fronterizo", "puerto_fronterizo"], "", true);
  const tempRequerida =
    getCsvValueLike(
      row,
      [
        "temp_requerida_nom_242_ssa1_2009",
        "temp_requerida_nom_242_ssa1_2009_art_6_4_8",
        "temp_requerida_nom_242_ssa1_2009_art_648",
        "temp_requerida",
        "temperatura_requerida",
        "temperatura_requerida_nom_242",
      ],
      "",
      true,
    ) || getCsvValueLike(row, ["temp_requerida", "temperatura_requerida", "nom_242"], "", true);
  const viajesAnuales2026 = getCsvValueLike(
    row,
    [
      "viajes_anuales_2026_est",
      "viajes_anuales_2026",
      "estimado_de_viajes_anuales_2026_est",
      "viajes_anuales_est",
      "viajes_anuales",
    ],
    "",
    true,
  );
  const duaMes = getCsvValueLike(row, ["dua_mes", "dua_mes_est", "duas_mes", "dua"], "", true);
  const volumenEstimado = getCsvValueLike(row, ["volume_estimado", "volumen_estimado", "volumen"], "", true);
  const volumenEstimadoFinal = volumenEstimado || viajesAnuales2026;
  const riesgoLogisticoCsv = getCsvValueLike(
    row,
    ["riesgo_logistico", "riesgo_logistico_operativo", "riesgo_logistico_empresa", "riesgo"],
    "",
    true,
  );
  const ventasAnuales = getCsvValueLike(
    row,
    ["ventas_anuales_estimadas", "ventas_estimadas", "ventas_anuales"],
    "",
    true,
  );
  const retencion = getCsvValueLike(row, ["retencion"], "", true);
  const tempFromAnyField = extractTempFromAnyField(row);
  const tempRequeridaFinal = stripNom242Tag(
    tempRequerida || tempFromAnyField || inferTempRequerida(productos, actividad),
  );
  const gamMatch = isGamCompanyName(empresa);
  const ubicacionFinal = gamMatch ? "Mazatlán, Sinaloa." : ubicacion;
  const rutaTerrestreFinal = gamMatch ? "Sinaloa -> Nogales -> EE.UU." : rutaTerrestre;
  const cruceFronterizoFinal = gamMatch ? "Nogales, AZ." : cruceFronterizo;
  const coord = inferCoordsBySede(ubicacionFinal);

  return {
    empresa,
    actividad,
    productos,
    certificaciones,
    frecuenciaSensor,
    ubicacion: ubicacionFinal,
    aniosOperacion,
    alcance,
    destinoFinal,
    paginaWeb,
    telefono,
    email,
    linkedin,
    rutaTerrestre: rutaTerrestreFinal,
    rutaMaritima,
    cruceFronterizo: cruceFronterizoFinal,
    volumenEstimado: volumenEstimadoFinal,
    viajesAnuales2026,
    duaMes,
    riesgoLogisticoCsv,
    ventasAnuales,
    retencion,
    tempRequerida: tempRequeridaFinal,
    contacto: buildContactoLabel(paginaWeb, telefono, email, linkedin),
    contactoLink: firstLinkValue(paginaWeb, telefono, email, linkedin),
    especialidad: productos,
    relevancia: inferRelevancia(alcance),
    lat: coord.lat,
    lng: coord.lng,
  };
}

function isGamCompanyName(name = "") {
  const key = normalizeHeader(name || "");
  return key.includes("grupo_acuicola_mexicano") || key === "gam" || key.includes("_gam");
}

function getCsvValueLike(row, aliases, fallback = NO_INFO, forContact = false) {
  const rowKeys = Object.keys(row || {});
  for (const alias of aliases) {
    const aliasNorm = normalizeHeader(alias);
    const key = rowKeys.find((k) => {
      const keyNorm = normalizeHeader(k);
      return keyNorm === aliasNorm || keyNorm.startsWith(aliasNorm);
    });
    if (!key) continue;
    const cleaned = cleanCell(row[key], { forContact });
    if (cleaned) return cleaned;
  }
  return fallback;
}

function inferTempRequerida(productos = "", actividad = "") {
  const text = normalizeGeoKey(`${productos} ${actividad}`);
  if (text.includes("CONSERVA")) {
    return "Conserva: Temp. ambiente | Congelado: <= -18 C";
  }
  if (text.includes("OSTION") || text.includes("ALMEJA") || text.includes("MEJILLON") || text.includes("VIVO")) {
    return "Vivo/Refrigerado: <= 7 C | Congelado: <= -18 C";
  }
  return "Fresco: <= 4 C | Congelado: <= -18 C";
}

function extractTempFromAnyField(row) {
  const values = Object.values(row || {})
    .map((v) => String(v || "").trim())
    .filter(Boolean);

  const byNom = values.find((value) => /NOM[\s-]?242/i.test(value));
  if (byNom) return byNom;

  const byTempPattern = values.find((value) => {
    const hasContext = /(CONGELADO|FRESCO|PASTEURIZADO|REFRIGERADO|VIVO|TEMP\.?\s*AMBIENTE)/i.test(value);
    const hasTempMetric = /(°\s*C|\b-?\d+\s*°?\s*C\b|[-−]\s*18\s*°?\s*C|(?:<=|≤)\s*-?\d+\s*°?\s*C)/i.test(value);
    return hasContext && hasTempMetric;
  });

  return byTempPattern || "";
}

function stripNom242Tag(text) {
  return String(text || "")
    .replace(/\(\s*NOM[\s-]?242[^)]*\)/gi, "")
    .replace(/\n\s*NOM[\s-]?242[^\n]*/gi, "")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function parseCsvText(text) {
  const normalized = String(text || "").replace(/^\uFEFF/, "").replace(/\r\n/g, "\n");
  const sampleLine = normalized
    .split("\n")
    .map((line) => line.trim())
    .find((line) => line.length > 0);
  const delimiter = detectCsvDelimiter(sampleLine || "");
  const rows = [];
  let row = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < normalized.length; i += 1) {
    const ch = normalized[i];
    const next = normalized[i + 1];
    if (ch === '"' && inQuotes && next === '"') {
      current += '"';
      i += 1;
      continue;
    }
    if (ch === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (ch === delimiter && !inQuotes) {
      row.push(current.trim());
      current = "";
      continue;
    }
    if (ch === "\n" && !inQuotes) {
      row.push(current.trim());
      current = "";
      if (row.some((cell) => cell !== "")) rows.push(row);
      row = [];
      continue;
    }
    current += ch;
  }
  row.push(current.trim());
  if (row.some((cell) => cell !== "")) rows.push(row);
  return rows;
}

function readCsvRows(csvText) {
  const parsed = parseCsvText(csvText);
  if (parsed.length) return parsed;

  try {
    if (typeof XLSX !== "undefined") {
      const wb = XLSX.read(String(csvText || ""), { type: "string" });
      const sheetName = wb.SheetNames?.[0];
      const sheet = sheetName ? wb.Sheets[sheetName] : null;
      if (sheet) {
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
        if (Array.isArray(rows) && rows.length) {
          return rows.map((row) => row.map((cell) => String(cell ?? "").trim()));
        }
      }
    }
  } catch (error) {
    console.error("readCsvRows fallback parseCsvText:", error);
  }
  return [];
}

async function fetchTextFromCandidates(candidates) {
  const unique = Array.from(new Set((candidates || []).filter(Boolean)));
  for (const path of unique) {
    try {
      const response = await fetch(path, { cache: "no-store" });
      if (!response.ok) continue;
      const text = await response.text();
      if (String(text || "").trim()) return text;
    } catch (error) {
      console.error(`No se pudo leer: ${path}`, error);
    }

    const xhrText = readTextWithXhr(path);
    if (String(xhrText || "").trim()) return xhrText;
  }
  return "";
}

function readTextWithXhr(path) {
  try {
    if (typeof XMLHttpRequest === "undefined") return "";
    const xhr = new XMLHttpRequest();
    xhr.open("GET", path, false);
    xhr.send(null);
    const ok = (xhr.status >= 200 && xhr.status < 300) || xhr.status === 0;
    if (!ok) return "";
    return String(xhr.responseText || "");
  } catch (error) {
    return "";
  }
}

function detectCsvDelimiter(line) {
  const options = [",", ";", "\t", "|"];
  let best = ",";
  let bestCount = -1;
  options.forEach((delimiter) => {
    const count = countDelimiterOutsideQuotes(line, delimiter);
    if (count > bestCount) {
      best = delimiter;
      bestCount = count;
    }
  });
  return best;
}

function countDelimiterOutsideQuotes(line, delimiter) {
  let inQuotes = false;
  let count = 0;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    const next = line[i + 1];
    if (ch === '"' && inQuotes && next === '"') {
      i += 1;
      continue;
    }
    if (ch === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (ch === delimiter && !inQuotes) count += 1;
  }
  return count;
}

function cleanCell(value, { forContact = false } = {}) {
  const text = String(value || "").trim();
  if (!text) return forContact ? "" : NO_INFO;
  const compact = normalizeHeader(text).replace(/_/g, "");
  if (["nd", "na", "sindato", "nodisponible", "nodisponiblepublicamente"].includes(compact)) {
    return forContact ? "" : NO_INFO;
  }
  return text;
}

function getCsvValue(row, aliases, fallback = NO_INFO, forContact = false) {
  for (const alias of aliases) {
    if (!(alias in row)) continue;
    const cleaned = cleanCell(row[alias], { forContact });
    if (cleaned) return cleaned;
  }
  return fallback;
}

function normalizeHeader(text) {
  return String(text || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[()]/g, "")
    .replace(/[./]/g, " ")
    .replace(/[^a-zA-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toLowerCase();
}

function inferCoordsBySede(sede) {
  const normalizedSede = normalizeHeader(sede).replace(/_/g, " ");
  const match = sedeCoords
    .filter((item) => normalizedSede.includes(item.key))
    .sort((a, b) => b.key.length - a.key.length)[0];
  return match ? { lat: match.lat, lng: match.lng } : { lat: 23.6345, lng: -102.5528 };
}

function inferCoordsByCompetidor(item = {}) {
  const lat = Number(item?.lat);
  const lng = Number(item?.lng);
  if (Number.isFinite(lat) && Number.isFinite(lng)) return { lat, lng };

  const empresaKey = normalizeGeoKey(item?.empresa || "");
  const byName = competidorCoordsByName.find((entry) => empresaKey.includes(entry.key));
  if (byName) return { lat: byName.lat, lng: byName.lng };

  return inferCoordsBySede(`${item?.sede || ""} ${item?.ciudad || ""}`.trim());
}

function inferRelevancia(alcance) {
  const value = normalizeHeader(alcance).replace(/_/g, " ");
  if (value.includes("estados unidos") || value.includes("europa") || value.includes("asia")) {
    return "Alta";
  }
  if (value.includes("internacional") || value.includes("export")) {
    return "Media-Alta";
  }
  if (value.includes("nacional")) {
    return "Media";
  }
  return "Media";
}

function splitLinks(raw) {
  return String(raw || "")
    .split(/[;,]/)
    .map((item) => item.trim())
    .filter(Boolean)
    .filter((item) => /^https?:\/\//i.test(item));
}

function buildContactoLabel(paginaWeb, telefono, email, linkedin) {
  if (telefono) return telefono;
  if (email) return email;
  const web = splitLinks(paginaWeb)[0];
  if (web) return web;
  const li = splitLinks(linkedin)[0];
  if (li) return li;
  return "No disponible";
}

function firstLinkValue(paginaWeb, telefono, email, linkedin) {
  const web = splitLinks(paginaWeb)[0];
  if (web) return web;
  const phoneDigits = extractFirstPhoneDigits(telefono);
  if (phoneDigits) return `https://wa.me/${phoneDigits}`;
  if (email) return `mailto:${email}`;
  const li = splitLinks(linkedin)[0];
  if (li) return li;
  return "";
}

function extractFirstPhoneDigits(raw) {
  const matches = String(raw || "").match(/\+?\d[\d\s().-]{8,}\d/g) || [];
  if (!matches.length) return "";
  const digits = matches[0].replace(/\D/g, "");
  return digits.length >= 10 ? digits : "";
}

function getInfraTerrestresNodes() {
  return (state.infraCruces || [])
    .filter((item) => item && item.sourceType === "terrestre")
    .slice(0, 6)
    .map((item, idx) => ({
      ...item,
      key: `terrestre-${idx}`,
      typeLabel: "Cruce terrestre",
    }));
}

function initInfraKpi() {
  const select = document.getElementById("infraNodoSelect");
  const summary = document.getElementById("infraNodoSummary");
  const pitch = document.getElementById("infraNodoPitch");
  if (!select || !summary || !pitch) return;

  const terrestres = getInfraTerrestresNodes();
  const nodes = [...terrestres];

  if (!nodes.length) {
    select.innerHTML = "<option value=''>Sin datos de infraestructura</option>";
    summary.innerHTML = "";
    pitch.innerHTML = "";
    return;
  }

  const terrestreOptions = terrestres
    .map((node) => `<option value="${node.key}">${escapeHtml(node.nombre)}</option>`)
    .join("");
  select.innerHTML = `
    <optgroup label="Cruces terrestres">${terrestreOptions}</optgroup>
  `;

  const fallbackKey = nodes[0].key;
  const selectedKey = nodes.some((n) => n.key === state.infraSelectedNodeId) ? state.infraSelectedNodeId : fallbackKey;
  state.infraSelectedNodeId = selectedKey;
  select.value = selectedKey;

  const renderSelected = ({ syncMap = true, openPopup = true } = {}) => {
    state.infraSelectedNodeId = select.value || fallbackKey;
    const node = nodes.find((n) => n.key === state.infraSelectedNodeId) || nodes[0];
    if (!node) return;
    const riskLabel = infraRiskLabel(node.riesgoLevel);
    const cTpat = infraCtpatDisplay(node.cTpatActivo, node.sourceType);
    const ftlMensual = Number.isFinite(node.ftlMensual)
      ? Math.round(node.ftlMensual).toLocaleString("es-MX")
      : Number.isFinite(node.ftlAnual)
        ? Math.round(node.ftlAnual / 12).toLocaleString("es-MX")
        : "N/D";
    const tiempoFda = node.tiempoFda || NO_INFO;
    const especiesRaw = compactRouteText(node.especies || NO_INFO);
    const productoPrincipalCruce =
      especiesRaw === NO_INFO
        ? NO_INFO
        : especiesRaw
            .split(/\s*,\s*/)
            .map((item) => item.trim())
            .filter(Boolean)
            .join(" · ");
    const pitchValue = compactRouteText(node.pitch || NO_INFO);

    summary.innerHTML = `
      <div class="infra-summary-head">
        <h4>${escapeHtml(node.nombre)}</h4>
        <div class="infra-summary-badges">
          <span class="infra-node-type">${escapeHtml(node.typeLabel)}</span>
          <span class="infra-risk-chip ${infraRiskClass(node.riesgoLevel)}">${riskLabel}</span>
        </div>
      </div>
      <div class="infra-metric-grid">
        <div class="infra-metric-item">
          <strong class="infra-metric-value">${escapeHtml(ftlMensual)}</strong>
          <span class="infra-metric-label" title="FTL — Full Truck Load">FTL/mes</span>
        </div>
        <div class="infra-metric-item">
          <strong class="infra-metric-value">${escapeHtml(tiempoFda)}</strong>
          <span class="infra-metric-label">Tiempo inspección FDA</span>
        </div>
        <div class="infra-metric-item">
          <strong class="infra-ctpat-chip ${cTpat.className}">${cTpat.label}</strong>
          <span class="infra-metric-label" title="C-TPAT — Customs-Trade Partnership Against Terrorism">C-TPAT</span>
          <small class="infra-ctpat-note">Reduce cruce ~60%</small>
        </div>
      </div>
      <div class="infra-species-line">Producto principal que cruza: ${escapeHtml(productoPrincipalCruce)}</div>
    `;
    pitch.innerHTML = `
      <p class="infra-pitch-quote">${escapeHtml(pitchValue)}</p>
    `;

    if (syncMap) {
      focusInfraCruceMarker(node.key, { openPopup });
    }
  };

  select.onchange = () => renderSelected({ syncMap: true, openPopup: true });
  renderInfraCrucesMap(nodes, (nodeKey) => {
    if (!nodeKey) return;
    select.value = nodeKey;
    renderSelected({ syncMap: true, openPopup: true });
  });
  renderSelected({ syncMap: true, openPopup: true });
}

function buildInfraCrucePopupHtml(cruce) {
  const riskLabel = infraRiskLabel(cruce.riesgoLevel);
  const ftlMes = Number.isFinite(cruce.ftlMensual)
    ? `${Math.round(cruce.ftlMensual).toLocaleString("es-MX")} FTL/mes`
    : "No disponible";
  const tiempoFda = cruce.tiempoFda || NO_INFO;
  const prospectos = getProspectosByCruce(cruce?.nombre || "");
  return `
    <div class="infra-popup">
      <strong>${escapeHtml(cruce.nombre)}</strong><br/>
      <span>FTL mariscos/mes: ${escapeHtml(ftlMes)}</span><br/>
      <span>Tiempo cruce FDA: ${escapeHtml(tiempoFda)}</span><br/>
      <span>Riesgo: <span class="infra-risk-chip ${infraRiskClass(cruce.riesgoLevel)}">${riskLabel}</span></span><br/>
      <span>Prospectos: ${escapeHtml(prospectos)}</span>
    </div>
  `;
}

function getProspectosByCruce(cruceNombre = "") {
  const key = normalizeGeoKey(cruceNombre);
  if (key.includes("NOGALES")) {
    return "Grupo Pinsa · GAM";
  }
  if (key.includes("OTAY") || (key.includes("TIJUANA") && key.includes("SAN DIEGO"))) {
    return "Baja Aqua-Farms · Pacífico Aquaculture · Baja Shellfish Farms";
  }
  return "ninguno identificado";
}

function buildInfraMarker(level, isActive = false) {
  const toneClass = level === "CRITICO" ? "infra-marker-critical" : level === "ALTO" ? "infra-marker-high" : "infra-marker-moderate";
  const activeClass = isActive ? "infra-marker-active" : "";
  return L.divIcon({
    className: "infra-marker-wrap",
    html: `<span class="infra-marker-triangle ${toneClass} ${activeClass}"></span>`,
    iconSize: isActive ? [20, 20] : [16, 16],
    iconAnchor: isActive ? [10, 16] : [8, 14],
    popupAnchor: [0, -14],
  });
}

function setActiveInfraMarker(nodeKey = "") {
  state.infraCrucesMarkers.forEach((entry) => {
    if (!entry || !entry.marker) return;
    const isActive = entry.key === nodeKey;
    entry.marker.setIcon(buildInfraMarker(entry.riesgoLevel, isActive));
  });
}

function focusInfraCruceMarker(nodeKey, { openPopup = true, doZoom = true } = {}) {
  if (!state.infraCrucesMap) return;
  const entry = (state.infraCrucesMarkers || []).find((item) => item && item.key === nodeKey);
  if (!entry || !entry.marker) return;

  setActiveInfraMarker(nodeKey);
  const latlng = entry.marker.getLatLng();
  if (doZoom) {
    const zoom = Math.max(state.infraCrucesMap.getZoom(), 5.8);
    state.infraCrucesMap.flyTo(latlng, zoom, { duration: 0.45 });
  }
  if (openPopup) entry.marker.openPopup();
}

function renderInfraCrucesMap(cruces = [], onMarkerSelect = null) {
  const mapEl = document.getElementById("infraCrucesMap");
  if (!mapEl || typeof L === "undefined") return;

  if (state.infraCrucesMap) {
    state.infraCrucesMap.remove();
    state.infraCrucesMap = null;
    state.infraCrucesMarkers = [];
  }

  const borderFocusCenter = [29.2, -107.2];
  const borderFocusZoom = 5.1;
  const map = L.map(mapEl).setView(borderFocusCenter, borderFocusZoom);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);

  if (!cruces.length) {
    L.marker([23.6345, -102.5528]).addTo(map).bindPopup("Sin cruces terrestres en CSV");
    state.infraCrucesMap = map;
    return;
  }

  const bounds = [];
  cruces.forEach((cruce) => {
    const isActive = cruce.key === state.infraSelectedNodeId;
    const marker = L.marker([cruce.lat, cruce.lng], { icon: buildInfraMarker(cruce.riesgoLevel, isActive) })
      .addTo(map)
      .bindPopup(buildInfraCrucePopupHtml(cruce));
    marker.on("click", () => {
      if (typeof onMarkerSelect === "function") onMarkerSelect(cruce.key);
    });
    state.infraCrucesMarkers.push({
      key: cruce.key,
      riesgoLevel: cruce.riesgoLevel,
      marker,
    });
    bounds.push([cruce.lat, cruce.lng]);
  });

  map.fitBounds(bounds, { padding: [30, 30], maxZoom: 6.5 });
  if (map.getZoom() < 4.9) {
    map.setView(borderFocusCenter, 4.9);
  }
  state.infraCrucesMap = map;
  if (state.infraSelectedNodeId) {
    setActiveInfraMarker(state.infraSelectedNodeId);
  }
}

function renderCompetidoresKpi() {
  const grid = document.getElementById("competidoresGrid");
  const mapEl = document.getElementById("competidoresMap");
  if (!grid) return;
  try {
    if (state.competidoresMap) {
      state.competidoresMap.remove();
      state.competidoresMap = null;
    }
    if (mapEl) mapEl.innerHTML = "";

    const visibleCompetidores = getPanoramaCompetidores(competidoresData);
    if (!visibleCompetidores.length) {
      grid.innerHTML = '<article class="competidor-card"><div class="competidor-meta">Sin competidores en el CSV.</div></article>';
      return;
    }
    grid.innerHTML = visibleCompetidores
      .map(
        (item, idx) => `
      <article class="competidor-card">
        <h4>${item.empresa}</h4>
        <div class="competidor-meta"><strong>Tipo:</strong> ${item.tipo}</div>
        <div class="competidor-meta"><strong>Sede:</strong> ${item.sede || item.sedeEnMexico || item.ciudad || NO_INFO}</div>
        <div class="competidor-meta"><strong>Servicio principal:</strong> ${item.servicio || item.servicioPrincipal || NO_INFO}</div>
        <div class="competidor-meta"><strong>Sitio:</strong> ${renderWebsite(item.sitio)}</div>
        <div class="competidor-actions">
          <button class="competidor-btn" type="button" data-comp-toggle="${idx}" aria-expanded="false">
            Ver propuesta de valor
          </button>
          <button class="competidor-btn" type="button" data-comp-model-toggle="${idx}" aria-expanded="false">
            Ver modelo de negocio
          </button>
        </div>
        <div class="competidor-propuesta" data-comp-propuesta="${idx}" hidden>
          ${item.propuestaValor || NO_INFO}
        </div>
        <div class="competidor-propuesta" data-comp-model="${idx}" hidden>
          ${item.modeloNegocio || NO_INFO}
        </div>
      </article>
    `,
      )
      .join("");

    grid.querySelectorAll("[data-comp-toggle]").forEach((button) => {
      button.addEventListener("click", () => {
        const idx = button.getAttribute("data-comp-toggle");
        const target = grid.querySelector(`[data-comp-propuesta="${idx}"]`);
        if (!target) return;
        const isHidden = target.hasAttribute("hidden");
        if (isHidden) {
          target.removeAttribute("hidden");
          button.setAttribute("aria-expanded", "true");
          button.textContent = "Ocultar propuesta de valor";
        } else {
          target.setAttribute("hidden", "");
          button.setAttribute("aria-expanded", "false");
          button.textContent = "Ver propuesta de valor";
        }
      });
    });

    grid.querySelectorAll("[data-comp-model-toggle]").forEach((button) => {
      button.addEventListener("click", () => {
        const idx = button.getAttribute("data-comp-model-toggle");
        const target = grid.querySelector(`[data-comp-model="${idx}"]`);
        if (!target) return;
        const isHidden = target.hasAttribute("hidden");
        if (isHidden) {
          target.removeAttribute("hidden");
          button.setAttribute("aria-expanded", "true");
          button.textContent = "Ocultar modelo de negocio";
        } else {
          target.setAttribute("hidden", "");
          button.setAttribute("aria-expanded", "false");
          button.textContent = "Ver modelo de negocio";
        }
      });
    });

    initCompetidoresMap(visibleCompetidores);
  } catch (error) {
    console.error("renderCompetidoresKpi fallback:", error);
    renderCompetidoresEmergencyFallback();
  }
}

const PANORAMA_COMPETIDORES_ORDER = ["SENSORGO", "SYCOD", "SENSITECH", "REDGPS"];

function getPanoramaCompetidores(rows = []) {
  const mapped = (rows || []).map((item) => ({
    ...item,
    __key: normalizeGeoKey(item?.empresa || ""),
  }));
  const filtered = mapped.filter((item) => PANORAMA_COMPETIDORES_ORDER.some((needle) => item.__key.includes(needle)));
  filtered.sort((a, b) => {
    const ai = PANORAMA_COMPETIDORES_ORDER.findIndex((needle) => a.__key.includes(needle));
    const bi = PANORAMA_COMPETIDORES_ORDER.findIndex((needle) => b.__key.includes(needle));
    return ai - bi;
  });
  return filtered.map(({ __key, ...item }) => item);
}

function getCompetidorTipoColor(tipo = "") {
  const normalized = normalizeGeoKey(tipo);
  if (normalized.includes("INTERNACIONAL")) return "#1a6b3a";
  if (normalized.includes("MEXICANA")) return "#2e8f4f";
  if (normalized.includes("DISTRIBUIDOR")) return "#8fcda2";
  return "#2e8f4f";
}

function getCompetidorEmpresaColor(empresa = "", fallbackIndex = 0) {
  const palette = [
    "#1a6b3a",
    "#0f4c81",
    "#c2410c",
    "#7c3aed",
    "#b91c1c",
    "#0f766e",
    "#1d4ed8",
    "#a16207",
    "#be185d",
    "#334155",
  ];
  return palette[Math.abs(Number(fallbackIndex) || 0) % palette.length];
}

function buildCompetidorColorMap(rows = []) {
  const uniqueKeys = Array.from(
    new Set(
      (rows || [])
        .map((item) => normalizeGeoKey(item?.empresa || item?.sede || item?.tipo || ""))
        .filter(Boolean),
    ),
  ).sort((a, b) => a.localeCompare(b));

  const colorByKey = new Map();
  uniqueKeys.forEach((key, idx) => {
    colorByKey.set(key, getCompetidorEmpresaColor(key, idx));
  });
  return colorByKey;
}

function initCompetidoresMap(rows = competidoresData) {
  const mapEl = document.getElementById("competidoresMap");
  if (!mapEl || typeof L === "undefined" || !rows.length) return;

  if (state.competidoresMap) {
    state.competidoresMap.remove();
    state.competidoresMap = null;
  }

  const map = L.map(mapEl, { zoomControl: false }).setView([23.5, -102.5], 4.7);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);

  const bounds = [];
  const coordCount = new Map();
  const colorByEmpresa = buildCompetidorColorMap(rows);

  rows.forEach((item, idx) => {
    const base = inferCoordsByCompetidor(item);
    const key = `${base.lat.toFixed(3)},${base.lng.toFixed(3)}`;
    const seen = coordCount.get(key) || 0;
    coordCount.set(key, seen + 1);

    const jitterStep = 0.022;
    const angle = seen * 2.2;
    const lat = base.lat + (seen ? Math.sin(angle) * jitterStep : 0);
    const lng = base.lng + (seen ? Math.cos(angle) * jitterStep : 0);
    const companyKey = normalizeGeoKey(item.empresa || item.sede || item.tipo || "");
    const color = colorByEmpresa.get(companyKey) || getCompetidorEmpresaColor(companyKey, idx);

    const popup = `
      <div class="infra-popup">
        <strong>${escapeHtml(compactRouteText(item.empresa || "Competidor"))}</strong><br/>
        <span><strong>Tipo:</strong> ${escapeHtml(compactRouteText(item.tipo || NO_INFO))}</span><br/>
        <span><strong>Presencia:</strong> ${escapeHtml(compactRouteText(item.sede || NO_INFO))}</span><br/>
        <span><strong>Servicio:</strong> ${escapeHtml(compactRouteText(item.servicio || NO_INFO))}</span>
      </div>
    `;

    const marker = L.circleMarker([lat, lng], {
      radius: 8,
      color,
      fillColor: color,
      fillOpacity: 0.92,
      weight: 2,
    })
      .addTo(map)
      .bindPopup(popup);

    bounds.push(marker.getLatLng());
  });

  if (bounds.length) {
    map.fitBounds(bounds, { padding: [44, 44], maxZoom: 6 });
  }

  state.competidoresMap = map;
}

function buildEmpresasHash(rows = []) {
  const normalizedRows = (rows || []).map((row) => {
    const ordered = {};
    Object.keys(row || {})
      .sort()
      .forEach((key) => {
        if (key === "lat" || key === "lng") return;
        ordered[key] = row[key];
      });
    return ordered;
  });
  return JSON.stringify(normalizedRows);
}

function getDefaultRiesgoEmpresaIndex() {
  if (!empresasData.length) return 0;
  const grupoPinsaIdx = empresasData.findIndex((item) =>
    normalizeGeoKey(item?.empresa || "").includes("GRUPO PINSA"),
  );
  return grupoPinsaIdx >= 0 ? grupoPinsaIdx : 0;
}

function syncRiesgoEmpresaOptions(forceDefault = false) {
  const select = document.getElementById("riesgoEmpresaSelect");
  if (!select) return;
  if (!empresasData.length) {
    select.innerHTML = "<option value='0'>Sin empresas</option>";
    return;
  }

  const defaultIndex = getDefaultRiesgoEmpresaIndex();
  const prevIndex = Number.parseInt(select.value, 10);
  select.innerHTML = empresasData
    .map((e, i) => `<option value="${i}">${e.empresa}</option>`)
    .join("");
  const safeIndex = forceDefault
    ? defaultIndex
    : Number.isFinite(prevIndex)
      ? Math.max(0, Math.min(prevIndex, empresasData.length - 1))
      : defaultIndex;
  select.value = String(safeIndex);
}

async function refreshEmpresasDataFromCsv() {
  const prevHash = state.empresasHash || "";
  const prevSource = state.empresasSource || "fallback";
  await loadEmpresasData();
  const nextHash = buildEmpresasHash(empresasData);
  const nextSource = state.empresasSource || "fallback";
  if (nextHash === prevHash && nextSource === prevSource) return;

  state.empresasHash = nextHash;
  renderEmpresas();
  renderPropuestaTab();
  renderClustering();
  syncRiesgoEmpresaOptions();

  const select = document.getElementById("riesgoEmpresaSelect");
  if (select && empresasData.length) {
    updateRiesgosByEmpresa(Number(select.value || 0));
  }
}

function initRiesgos() {
  const select = document.getElementById("riesgoEmpresaSelect");
  const terrestreSelect = document.getElementById("riesgoRutaTerrestreSelect");
  const maritimaSelect = document.getElementById("riesgoRutaMaritimaSelect");
  const verUbicacionBtn = document.getElementById("riesgoVerUbicacionBtn");
  const verProductosBtn = document.getElementById("riesgoVerProductosBtn");
  const verCertificacionesBtn = document.getElementById("riesgoVerCertificacionesBtn");
  const verAduanaBtn = document.getElementById("riesgoVerAduanaBtn");
  const ubicacionInfo = document.getElementById("riesgoUbicacionInfo");
  const productosInfo = document.getElementById("riesgoProductosInfo");
  const certificacionesInfo = document.getElementById("riesgoCertificacionesInfo");
  const aduanaInfo = document.getElementById("riesgoAduanaInfo");
  if (!select) return;

  const hasValidIndex = Number.isFinite(Number.parseInt(select.value, 10));
  syncRiesgoEmpresaOptions(!hasValidIndex);
  if (!empresasData.length) return;
  syncRiesgoMesOperacionLabel();

  if (!state.riesgosListenersBound) {
    select.addEventListener("change", () => {
      updateRiesgosByEmpresa(Number(select.value));
    });

    if (terrestreSelect) {
      terrestreSelect.addEventListener("change", () => {
        const idx = Number(select.value || 0);
        const empresa = empresasData[idx] || empresasData[0];
        if (empresa) state.riesgoSelectedTerrestre[empresa.empresa] = terrestreSelect.value || "";
        updateRiesgosByEmpresa(idx);
      });
    }

    if (maritimaSelect) {
      maritimaSelect.addEventListener("change", () => {
        const idx = Number(select.value || 0);
        const empresa = empresasData[idx] || empresasData[0];
        if (empresa) state.riesgoSelectedMaritima[empresa.empresa] = maritimaSelect.value || "";
        updateRiesgosByEmpresa(idx);
      });
    }

    if (verUbicacionBtn && ubicacionInfo) {
      verUbicacionBtn.addEventListener("click", () => {
        const isHidden = ubicacionInfo.hasAttribute("hidden");
        if (isHidden) {
          ubicacionInfo.removeAttribute("hidden");
          verUbicacionBtn.setAttribute("aria-expanded", "true");
          verUbicacionBtn.textContent = "Ocultar ubicación";
        } else {
          ubicacionInfo.setAttribute("hidden", "");
          verUbicacionBtn.setAttribute("aria-expanded", "false");
          verUbicacionBtn.textContent = "Ver ubicación";
        }
      });
    }

    if (verProductosBtn && productosInfo) {
      verProductosBtn.addEventListener("click", () => {
        const isHidden = productosInfo.hasAttribute("hidden");
        if (isHidden) {
          productosInfo.removeAttribute("hidden");
          verProductosBtn.setAttribute("aria-expanded", "true");
          verProductosBtn.textContent = "Ocultar productos";
        } else {
          productosInfo.setAttribute("hidden", "");
          verProductosBtn.setAttribute("aria-expanded", "false");
          verProductosBtn.textContent = "Ver productos";
        }
      });
    }

    if (verCertificacionesBtn && certificacionesInfo) {
      verCertificacionesBtn.addEventListener("click", () => {
        const isHidden = certificacionesInfo.hasAttribute("hidden");
        if (isHidden) {
          certificacionesInfo.removeAttribute("hidden");
          verCertificacionesBtn.setAttribute("aria-expanded", "true");
          verCertificacionesBtn.textContent = "Ocultar certificaciones";
        } else {
          certificacionesInfo.setAttribute("hidden", "");
          verCertificacionesBtn.setAttribute("aria-expanded", "false");
          verCertificacionesBtn.textContent = "Ver certificaciones";
        }
      });
    }

    if (verAduanaBtn && aduanaInfo) {
      verAduanaBtn.addEventListener("click", () => {
        const isHidden = aduanaInfo.hasAttribute("hidden");
        if (isHidden) {
          aduanaInfo.removeAttribute("hidden");
          verAduanaBtn.setAttribute("aria-expanded", "true");
          verAduanaBtn.textContent = "Ocultar aduana cruce";
        } else {
          aduanaInfo.setAttribute("hidden", "");
          verAduanaBtn.setAttribute("aria-expanded", "false");
          verAduanaBtn.textContent = "Ver aduana cruce";
        }
      });
    }

    state.riesgosListenersBound = true;
  }

  updateRiesgosByEmpresa(Number(select.value || 0));
}

function splitRouteOptions(rawRoute, fallback = "") {
  const lines = String(rawRoute || "")
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean);

  const expanded = [];
  lines.forEach((line) => {
    const pipeChunks = line
      .split(/\s+\|\s+/)
      .map((chunk) => chunk.trim())
      .filter(Boolean);

    pipeChunks.forEach((chunk) => {
      const semicolonChunks = chunk
        .split(/\s*;\s*/)
        .map((part) => part.trim())
        .filter(Boolean);

      semicolonChunks.forEach((part) => {
        const slashParts = part
          .split(/\s+\/\s+/)
          .map((item) => item.trim())
          .filter(Boolean);

        const arrowParts = slashParts.filter((item) => item.includes("->") || item.includes("→"));
        if (slashParts.length > 1 && arrowParts.length === slashParts.length && arrowParts.length >= 2) {
          expanded.push(...slashParts);
          return;
        }
        expanded.push(part);
      });
    });
  });

  const unique = Array.from(new Set(expanded.map((item) => compactRouteText(item)).filter(Boolean)));
  const fallbackText = String(fallback || "").trim();
  if (!unique.length && fallbackText) unique.push(fallbackText);
  return unique;
}

function routeOptionLabel(route, idx) {
  const compact = compactRouteText(route);
  const short = compact.length > 90 ? `${compact.slice(0, 90)}...` : compact;
  return `Ruta ${idx + 1}: ${short}`;
}

function humanizeRouteText(route = "") {
  let text = compactRouteText(route);
  const replacements = [
    [/\bLAX\b/gi, "Los Angeles, California"],
    [/\bLA\b/g, "Los Angeles, California"],
    [/\bCDMX\b/gi, "Ciudad de México"],
    [/\bGDL\b/gi, "Guadalajara"],
    [/\bMTY\b/gi, "Monterrey"],
    [/\bSF\b/gi, "San Francisco, California"],
  ];
  replacements.forEach(([pattern, value]) => {
    text = text.replace(pattern, value);
  });
  return text;
}

function mergeRouteWithDestino(routeText = "", destinoText = "", mode = "terrestre") {
  const route = compactRouteText(routeText);
  const destino = compactRouteText(destinoText);
  if (!hasMeaningfulRoute(destino)) return route;

  const routeNorm = normalizeGeoKey(route);
  const destinoNorm = normalizeGeoKey(destino);
  if (!destinoNorm || routeNorm.includes(destinoNorm)) return route;

  const destinoPoints = extractGeoWaypointsFromRoute(destino, mode);
  if (!destinoPoints.length) return route;

  const routePoints = extractGeoWaypointsFromRoute(route, mode);
  if (routePoints.length) {
    const lastRouteLabel = routePoints[routePoints.length - 1].label;
    const overlapIdx = destinoPoints.findIndex((point) => point.label === lastRouteLabel);
    if (overlapIdx >= 0) {
      const tail = destinoPoints.slice(overlapIdx + 1).map((point) => point.label);
      if (!tail.length) return route;
      return `${route} -> ${tail.join(" -> ")}`;
    }
  }

  const destinoTrail = destinoPoints.map((point) => point.label).join(" -> ");
  if (!destinoTrail) return route;
  if (!route) return destinoTrail;
  return `${route} -> ${destinoTrail}`;
}

function hasUsDestinationSignal(text = "") {
  const normalized = normalizeGeoKey(text);
  if (!normalized) return false;
  const keys = [
    "LOS ANGELES",
    "SAN FRANCISCO",
    "SAN DIEGO",
    "LAREDO",
    "HOUSTON",
    "PHOENIX",
    "MIAMI",
    "EL PASO",
    "MCALLEN",
    "BROWNSVILLE",
    "ESTADOS UNIDOS",
    "EE UU",
    "USA",
  ];
  return keys.some((key) => normalized.includes(key));
}

function shouldExtendTerrestreRoute(routeText = "", destinoText = "") {
  const route = compactRouteText(routeText);
  const destino = compactRouteText(destinoText);
  if (!hasMeaningfulRoute(route) || !hasMeaningfulRoute(destino)) return false;
  if (hasUsDestinationSignal(route)) return false;
  const destinoPoints = extractGeoWaypointsFromRoute(destino, "terrestre");
  if (!destinoPoints.length) return false;
  const routePoints = extractGeoWaypointsFromRoute(route, "terrestre");
  if (routePoints.length >= 3) return false;
  return true;
}

function formatRouteAlternatives(routes = [], prefix = "Ruta") {
  if (!Array.isArray(routes) || !routes.length) return `${prefix}: No disponible`;
  if (routes.length === 1) return `${prefix}: ${humanizeRouteText(routes[0])}`;
  const joined = routes
    .map((route, idx) => `Ruta ${idx + 1}: ${humanizeRouteText(route)}`)
    .join(" | ");
  return `${prefix}s: ${joined}`;
}

function fillRouteSelect(selectEl, options, selectedRoute, emptyLabel = "No aplica") {
  if (!selectEl) return options.includes(selectedRoute) ? selectedRoute : options[0] || "";
  selectEl.innerHTML = "";

  if (!options.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = emptyLabel;
    selectEl.append(option);
    selectEl.disabled = true;
    return "";
  }

  options.forEach((route, idx) => {
    const option = document.createElement("option");
    option.value = route;
    option.textContent = routeOptionLabel(route, idx);
    selectEl.append(option);
  });
  selectEl.disabled = options.length <= 1;

  const selected = options.includes(selectedRoute) ? selectedRoute : options[0];
  selectEl.value = selected;
  return selected;
}

async function updateRiesgosByEmpresa(index) {
  if (!empresasData.length) return;
  const empresa = empresasData[index] || empresasData[0];
  const verUbicacionBtn = document.getElementById("riesgoVerUbicacionBtn");
  const verProductosBtn = document.getElementById("riesgoVerProductosBtn");
  const verCertificacionesBtn = document.getElementById("riesgoVerCertificacionesBtn");
  const verAduanaBtn = document.getElementById("riesgoVerAduanaBtn");
  const ubicacionInfo = document.getElementById("riesgoUbicacionInfo");
  const productosInfo = document.getElementById("riesgoProductosInfo");
  const certificacionesInfo = document.getElementById("riesgoCertificacionesInfo");
  const aduanaInfo = document.getElementById("riesgoAduanaInfo");
  const productosEmpresa = compactRouteText(empresa.productos || empresa.especialidad || NO_INFO);
  const certificacionesEmpresa = compactRouteText(empresa.certificaciones || NO_INFO);
  const ubicacionEmpresa = compactRouteText(empresa.ubicacion || NO_INFO);
  if (ubicacionInfo) {
    ubicacionInfo.textContent = `Ubicación: ${ubicacionEmpresa}`;
    ubicacionInfo.setAttribute("hidden", "");
  }
  if (productosInfo) {
    productosInfo.textContent = `Productos: ${productosEmpresa}`;
    productosInfo.setAttribute("hidden", "");
  }
  if (certificacionesInfo) {
    certificacionesInfo.textContent = `Certificaciones: ${certificacionesEmpresa}`;
    certificacionesInfo.setAttribute("hidden", "");
  }
  if (aduanaInfo) {
    aduanaInfo.textContent = "Aduana cruce: No disponible";
    aduanaInfo.setAttribute("hidden", "");
  }
  if (verUbicacionBtn) {
    verUbicacionBtn.setAttribute("aria-expanded", "false");
    verUbicacionBtn.textContent = "Ver ubicación";
  }
  if (verProductosBtn) {
    verProductosBtn.setAttribute("aria-expanded", "false");
    verProductosBtn.textContent = "Ver productos";
  }
  if (verCertificacionesBtn) {
    verCertificacionesBtn.setAttribute("aria-expanded", "false");
    verCertificacionesBtn.textContent = "Ver certificaciones";
  }
  if (verAduanaBtn) {
    verAduanaBtn.setAttribute("aria-expanded", "false");
    verAduanaBtn.textContent = "Ver aduana cruce";
  }
  const empresaKey = empresa.empresa;
  const rutas = resolveEmpresaRutas(empresa);
  const rutaTerrestre = rutas.terrestre;
  const rutaOceanica = rutas.oceanica;
  const rutaTerrestreDetalleAll = getEmpresaRutaTerrestreRaw(empresa, rutaTerrestre.nombre);
  const rutaMaritimaDetalleAll = getEmpresaRutaMaritimaRaw(empresa, rutaOceanica.nombre);
  const destinoMercado = compactRouteText(empresa.destinoFinal || empresa.alcance || "");
  const destinoMercadoGeo = extractGeoWaypointsFromRoute(destinoMercado, "terrestre");
  const destinoTerrestreFallback = destinoMercadoGeo.length ? destinoMercado : "";
  const destinoTerrestreGeo = destinoMercadoGeo;

  const meteoNivel = document.getElementById("riesgoMeteoNivel");
  const meteoTexto = document.getElementById("riesgoMeteoTexto");
  const frioNivel = document.getElementById("riesgoFrioNivel");
  const frioTexto = document.getElementById("riesgoFrioTexto");
  const logNivel = document.getElementById("riesgoLogNivel");
  const logTexto = document.getElementById("riesgoLogTexto");
  const aduanaNivel = document.getElementById("riesgoAduanaNivel");
  const aduanaTexto = document.getElementById("riesgoAduanaTexto");
  const regNivel = document.getElementById("riesgoRegNivel");
  const regTexto = document.getElementById("riesgoRegTexto");
  const pronosticoTexto = document.getElementById("riesgoPronosticoTexto");
  const rutaTerrestreTexto = document.getElementById("riesgoRutaTerrestreTexto");
  const rutaMaritimaTexto = document.getElementById("riesgoRutaMaritimaTexto");
  const riesgoPitchTexto = document.getElementById("riesgoPitchTexto");
  const rutaTerrestreSelect = document.getElementById("riesgoRutaTerrestreSelect");
  const rutaMaritimaSelect = document.getElementById("riesgoRutaMaritimaSelect");
  const oleajeTexto = document.getElementById("riesgoOleajeTexto");

  const terrestres = splitRouteOptions(rutaTerrestreDetalleAll, rutaTerrestre.nombre);
  const maritimas = splitRouteOptions(rutaMaritimaDetalleAll, rutaOceanica.nombre).filter((route) =>
    hasMeaningfulRoute(route),
  );

  const rutaTerrestreDetalle = fillRouteSelect(
    rutaTerrestreSelect,
    terrestres,
    state.riesgoSelectedTerrestre[empresaKey] || "",
    "Sin rutas terrestres",
  );
  state.riesgoSelectedTerrestre[empresaKey] = rutaTerrestreDetalle;
  const shouldExtendSelectedRoute = shouldExtendTerrestreRoute(rutaTerrestreDetalle, destinoTerrestreFallback);
  const rutaTerrestreConDestino = shouldExtendSelectedRoute
    ? mergeRouteWithDestino(rutaTerrestreDetalle, destinoTerrestreFallback, "terrestre")
    : rutaTerrestreDetalle;
  const rutaTerrestreMapaRaw = hasMeaningfulRoute(rutaTerrestreDetalle)
    ? rutaTerrestreDetalle
    : rutaTerrestreConDestino;

  const rutaMaritimaDetalle = fillRouteSelect(
    rutaMaritimaSelect,
    maritimas,
    state.riesgoSelectedMaritima[empresaKey] || "",
    "No aplica",
  );
  state.riesgoSelectedMaritima[empresaKey] = rutaMaritimaDetalle;

  const hasMaritimaRutaCsv = maritimas.length > 0 && hasMeaningfulRoute(rutaMaritimaDetalle);
  if (rutaTerrestreTexto) {
    const baseRoutes = terrestres.length ? terrestres : [rutaTerrestreDetalle || rutaTerrestre.nombre || ""];
    const routes = Array.from(new Set(baseRoutes.map((route) => compactRouteText(route)).filter(Boolean)));
    const routeText = formatRouteAlternatives(routes, "Ruta terrestre");
    const shouldShowDestinoFinal = baseRoutes.some((route) => shouldExtendTerrestreRoute(route, destinoTerrestreFallback));
    if (
      shouldShowDestinoFinal &&
      destinoTerrestreGeo.length &&
      !normalizeGeoKey(routeText).includes(normalizeGeoKey(destinoTerrestreFallback))
    ) {
      rutaTerrestreTexto.textContent = `${routeText} | Destino final: ${humanizeRouteText(destinoTerrestreFallback)}`;
    } else {
      rutaTerrestreTexto.textContent = routeText;
    }
  }
  if (rutaMaritimaTexto) {
    if (hasMaritimaRutaCsv) {
      const routes = maritimas.length ? maritimas : [rutaMaritimaDetalle || rutaOceanica.nombre || ""];
      rutaMaritimaTexto.textContent = formatRouteAlternatives(routes, "Ruta marítima");
    } else {
      rutaMaritimaTexto.textContent = "Ruta marítima: No aplica";
    }
  }
  const rutaOceanicaCsv = detectOceanicaByText(normalizeGeoKey(rutaMaritimaDetalle)) || rutaOceanica;
  const rutaTerrestreCsv =
    detectTerrestreByText(normalizeGeoKey(`${empresa.cruceFronterizo || ""} ${rutaTerrestreMapaRaw}`)) ||
    rutaTerrestre;
  const cruceInfo =
    findCruceByEmpresa(
      empresa,
      `${rutaTerrestreMapaRaw || ""} | ${rutaTerrestreCsv?.nombre || ""} | ${rutaTerrestre?.nombre || ""}`,
    ) ||
    findCruceByRutaReferencia(
      `${rutaTerrestreCsv?.nombre || ""} | ${empresa?.cruceFronterizo || ""} | ${rutaTerrestreDetalle || ""}`,
    );
  if (aduanaInfo) {
    const aduanaCruce = compactRouteText(cruceInfo?.nombre || empresa?.cruceFronterizo || rutaTerrestreCsv?.nombre || NO_INFO);
    aduanaInfo.textContent = `Aduana cruce: ${aduanaCruce}`;
  }
  const mesOperacion = getSelectedRiesgoMes();
  const tempRequeridaCsv = stripNom242Tag(
    empresa.tempRequerida || inferTempRequerida(empresa.productos, empresa.actividad),
  );
  const estadosTransitoProxy = resolveTransitStateKeys(empresa, rutaTerrestreCsv.nombre, rutaTerrestreMapaRaw);
  let frioProxy = buildColdProxyFallback(estadosTransitoProxy, mesOperacion, tempRequeridaCsv);
  try {
    frioProxy = await calcColdChainProxyRisk(estadosTransitoProxy, mesOperacion, tempRequeridaCsv);
  } catch (proxyError) {
    console.error("cold proxy error:", proxyError);
  }

  const preLogRisk = parseEmpresaLogRiskCsv(empresa.riesgoLogisticoCsv || "") || calcLogRiskFallback("medio", rutaTerrestreCsv.nombre);
  const preAduanaRisk = buildAduanaRiskCard(cruceInfo);
  const preRegRisk = calcRegulatoryRisk(empresa, preLogRisk.level, cruceInfo);
  setRiskPill(meteoNivel, "Por actualizar", "medio");
  setRiskPill(frioNivel, titleCase(frioProxy.level), frioProxy.level);
  setRiskPill(logNivel, preLogRisk.label, preLogRisk.level);
  setAduanaRiskPill(aduanaNivel, preAduanaRisk.label, preAduanaRisk.level);
  setRiskPill(regNivel, preRegRisk.label, preRegRisk.level);
  if (meteoTexto) meteoTexto.textContent = "Actualizando clima en tiempo real...";
  if (frioTexto) {
    frioTexto.innerHTML = buildColdChainProxyText({
      tempRequerida: tempRequeridaCsv,
      maxTemp: null,
      proxy: frioProxy,
      mesOperacion,
    });
  }
  if (logTexto) {
    logTexto.innerHTML = buildLogisticRiskFields({
      logRiskCard: preLogRisk,
      cruceInfo,
    });
  }
  if (aduanaTexto) aduanaTexto.innerHTML = preAduanaRisk.text;
  if (regTexto) regTexto.innerHTML = buildRegulatoryRiskFields(preRegRisk);
  renderRiesgoTotalBanner(
    calcTotalRiskScore({
      cadenaFria: frioProxy.level,
      logistico: preLogRisk.level,
      regulatorio: preRegRisk.level,
      meteorologico: "medio",
    }),
  );

  try {
    const clima = await fetchOpenMeteo(empresa.lat, empresa.lng);
    const meteoRisk = calcMeteoRisk(clima);
    const frioRisk = calcColdChainRisk(clima);
    const frioNivelFinal = higherRiskLevel(frioRisk.level, frioProxy.level);

    setRiskPill(meteoNivel, meteoRisk.label, meteoRisk.level);
    setRiskPill(frioNivel, titleCase(frioNivelFinal), frioNivelFinal);

    const maxTemp = Math.max(...(clima.daily.temperature_2m_max || [0]));
    const maxViento = Math.max(...(clima.daily.wind_speed_10m_max || [0]));
    const maxLluvia = Math.max(...(clima.daily.precipitation_sum || [0]));

    if (meteoTexto) {
      meteoTexto.innerHTML = buildMeteoRiskText({
        maxTemp,
        maxViento,
        maxLluvia,
        level: meteoRisk.level,
      });
    }
    if (frioTexto) {
      frioTexto.innerHTML = buildColdChainProxyText({
        tempRequerida: tempRequeridaCsv,
        maxTemp,
        proxy: frioProxy,
        mesOperacion,
      });
    }
    if (pronosticoTexto) pronosticoTexto.textContent = `${empresa.empresa} (${empresa.ubicacion})`;
    renderForecast7d(clima);

    const portPoint = puertoCoords[rutaOceanicaCsv.nombre];
    let waveLevel = "medio";
    if (hasMaritimaRutaCsv && portPoint) {
      const marine = await fetchMarine(portPoint.lat, portPoint.lng);
      const wave = calcWaveRisk(marine);
      waveLevel = wave.level;
      if (oleajeTexto) {
        oleajeTexto.textContent = `Puerto: ${rutaOceanicaCsv.nombre}. Altura max de oleaje (48h): ${formatNumber(wave.maxWave, "m")} (${wave.label}).`;
      }
    } else if (hasMaritimaRutaCsv) {
      if (oleajeTexto) oleajeTexto.textContent = `Sin coordenadas marinas para ${rutaOceanicaCsv.nombre}.`;
    } else {
      if (oleajeTexto) oleajeTexto.textContent = "Ruta marítima no aplica para esta empresa.";
    }

    const logRiskRealtime = await calcLogRiskRealtime(
      meteoRisk.level,
      rutaTerrestreCsv.nombre,
      rutaTerrestreMapaRaw,
    );
    const logRiskCsv = parseEmpresaLogRiskCsv(empresa.riesgoLogisticoCsv || "");
    const logRiskMerged = mergeLogRiskSources(logRiskRealtime, logRiskCsv);
    const logRiskCard = logRiskCsv || logRiskMerged;
    const aduanaRisk = buildAduanaRiskCard(cruceInfo);
    const regRisk = calcRegulatoryRisk(empresa, logRiskCard.level, cruceInfo);
    setRiskPill(logNivel, logRiskCard.label, logRiskCard.level);
    if (logTexto) {
      logTexto.innerHTML = buildLogisticRiskFields({
        logRiskCard,
        cruceInfo,
      });
    }
    setAduanaRiskPill(aduanaNivel, aduanaRisk.label, aduanaRisk.level);
    if (aduanaTexto) aduanaTexto.innerHTML = aduanaRisk.text;
    setRiskPill(regNivel, regRisk.label, regRisk.level);
    if (regTexto) regTexto.innerHTML = buildRegulatoryRiskFields(regRisk);
    const totalRisk = calcTotalRiskScore({
      cadenaFria: frioNivelFinal,
      logistico: logRiskCard.level,
      regulatorio: regRisk.level,
      meteorologico: meteoRisk.level,
    });
    renderRiesgoTotalBanner(totalRisk);
    if (riesgoPitchTexto) {
      riesgoPitchTexto.textContent = buildRiesgoPitchText({
        empresa,
        mesOperacion,
        cruceInfo,
        frioProxy,
        logRiskCard,
        regRisk,
        meteoRisk,
        maxTemp,
      });
    }
    renderRiesgoRutaMap(empresa, rutaTerrestreCsv, logRiskMerged.level, rutaTerrestreMapaRaw);
    if (hasMaritimaRutaCsv) {
      renderRiesgoRutaMaritimaMap(empresa, rutaOceanicaCsv, waveLevel, rutaMaritimaDetalle);
    } else {
      renderRiesgoRutaMaritimaNoAplica();
    }
  } catch (error) {
    setRiskPill(meteoNivel, "No disponible", "medio");
    setRiskPill(frioNivel, titleCase(frioProxy.level), frioProxy.level);
    setRiskPill(logNivel, "No disponible", "medio");
    setRiskPill(regNivel, "No disponible", "medio");
    if (meteoTexto) meteoTexto.textContent = "No se pudo cargar Open-Meteo desde este entorno.";
    if (frioTexto) {
      frioTexto.innerHTML = buildColdChainProxyText({
        tempRequerida: tempRequeridaCsv,
        maxTemp: null,
        proxy: frioProxy,
        mesOperacion,
      });
    }
    const logRiskFallback = await calcLogRiskRealtime(
      "bajo",
      rutaTerrestreCsv.nombre,
      rutaTerrestreMapaRaw,
    );
    const logRiskCsv = parseEmpresaLogRiskCsv(empresa.riesgoLogisticoCsv || "");
    const logRiskMerged = mergeLogRiskSources(logRiskFallback, logRiskCsv);
    const logRiskCard = logRiskCsv || logRiskMerged;
    const aduanaRisk = buildAduanaRiskCard(cruceInfo);
    const regRisk = calcRegulatoryRisk(empresa, logRiskCard.level, cruceInfo);
    setRiskPill(logNivel, logRiskCard.label, logRiskCard.level);
    if (logTexto) {
      logTexto.innerHTML = buildLogisticRiskFields({
        logRiskCard,
        cruceInfo,
      });
    }
    setAduanaRiskPill(aduanaNivel, aduanaRisk.label, aduanaRisk.level);
    if (aduanaTexto) aduanaTexto.innerHTML = aduanaRisk.text;
    setRiskPill(regNivel, regRisk.label, regRisk.level);
    if (regTexto) regTexto.innerHTML = buildRegulatoryRiskFields(regRisk);
    const totalRisk = calcTotalRiskScore({
      cadenaFria: frioProxy.level,
      logistico: logRiskCard.level,
      regulatorio: regRisk.level,
      meteorologico: "medio",
    });
    renderRiesgoTotalBanner(totalRisk);
    if (riesgoPitchTexto) {
      riesgoPitchTexto.textContent = buildRiesgoPitchText({
        empresa,
        mesOperacion,
        cruceInfo,
        frioProxy,
        logRiskCard,
        regRisk,
        meteoRisk: { level: "medio" },
        maxTemp: frioProxy?.avgMaxTemp,
      });
    }
    if (pronosticoTexto) pronosticoTexto.textContent = "Sin pronóstico disponible.";
    renderForecast7d(null);
    renderRiesgoRutaMap(empresa, rutaTerrestreCsv, logRiskMerged.level, rutaTerrestreMapaRaw);
    if (hasMaritimaRutaCsv) {
      renderRiesgoRutaMaritimaMap(empresa, rutaOceanicaCsv, "medio", rutaMaritimaDetalle);
    } else {
      renderRiesgoRutaMaritimaNoAplica();
    }
    if (oleajeTexto) oleajeTexto.textContent = "Sin información de oleaje disponible.";
    console.error(error);
  }
}

async function fetchJsonWithTimeout(url, timeoutMs = 8000, sourceName = "API") {
  const controller = typeof AbortController !== "undefined" ? new AbortController() : null;
  const timeoutId = controller ? setTimeout(() => controller.abort(), timeoutMs) : null;
  try {
    const res = await fetch(url, {
      cache: "no-store",
      signal: controller ? controller.signal : undefined,
    });
    if (!res.ok) throw new Error(`${sourceName} no respondió.`);
    return res.json();
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new Error(`${sourceName} excedió tiempo de espera (${timeoutMs} ms).`);
    }
    throw error;
  } finally {
    if (timeoutId) clearTimeout(timeoutId);
  }
}

async function fetchOpenMeteo(lat, lng) {
  const url =
    `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lng}` +
    "&current=temperature_2m,wind_speed_10m,precipitation" +
    "&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,wind_speed_10m_max,weather_code&forecast_days=7&timezone=auto";
  return fetchJsonWithTimeout(url, RISK_API_TIMEOUT_MS, "Open-Meteo");
}

async function fetchMarine(lat, lng) {
  const url =
    `https://marine-api.open-meteo.com/v1/marine?latitude=${lat}&longitude=${lng}` +
    "&hourly=wave_height&forecast_days=2&timezone=auto";
  return fetchJsonWithTimeout(url, RISK_API_TIMEOUT_MS, "Marine API");
}

function calcMeteoRisk(clima) {
  const t = Math.max(...(clima.daily.temperature_2m_max || [0]));
  const v = Math.max(...(clima.daily.wind_speed_10m_max || [0]));
  const p = Math.max(...(clima.daily.precipitation_sum || [0]));
  if (t > 38 || v > 55 || p > 40) return { level: "alto", label: "Alto" };
  if (t > 32 || v > 40 || p > 20) return { level: "medio", label: "Medio" };
  return { level: "bajo", label: "Bajo" };
}

function toggleMeteoThresholds(button) {
  if (!button) return;
  const panel = button.nextElementSibling;
  if (!panel) return;
  const isOpen = !panel.hasAttribute("hidden");
  if (isOpen) {
    panel.setAttribute("hidden", "");
    button.setAttribute("aria-expanded", "false");
    button.textContent = "Ver umbrales aplicados";
  } else {
    panel.removeAttribute("hidden");
    button.setAttribute("aria-expanded", "true");
    button.textContent = "Ocultar umbrales";
  }
}

function buildMeteoRiskText({ maxTemp, maxViento, maxLluvia, level }) {
  const tempText = Number.isFinite(maxTemp) ? escapeHtml(formatNumber(maxTemp, "°C")) : "No disponible";
  const vientoText = Number.isFinite(maxViento) ? escapeHtml(formatNumber(maxViento, "km/h")) : "No disponible";
  const lluviaText = Number.isFinite(maxLluvia) ? escapeHtml(formatNumber(maxLluvia, "mm")) : "No disponible";

  const isBajo = level === "bajo";
  const isMedio = level === "medio";
  const isAlto = level === "alto";

  return `
    <div class="meteo-kpi-list">
      <div class="meteo-kpi-item">
        <span class="meteo-kpi-label">Temp máx (7 días)</span>
        <span class="meteo-kpi-value">${tempText}</span>
      </div>
      <div class="meteo-kpi-item">
        <span class="meteo-kpi-label">Viento máx (7 días)</span>
        <span class="meteo-kpi-value">${vientoText}</span>
      </div>
      <div class="meteo-kpi-item">
        <span class="meteo-kpi-label">Lluvia acum. (7 días)</span>
        <span class="meteo-kpi-value">${lluviaText}</span>
      </div>
    </div>
    <div class="cold-thresholds-wrap">
      <button
        type="button"
        class="cold-threshold-toggle"
        aria-expanded="false"
        onclick="toggleMeteoThresholds(this)"
      >
        Ver umbrales aplicados
      </button>
      <div class="cold-thresholds-panel" hidden>
        <div class="cold-thresholds-head">
          Regla: se evalúa el máximo de los próximos 7 días para temperatura, viento y lluvia.
        </div>
        <div class="cold-thresholds-list">
          <div class="cold-threshold-item ${isBajo ? "is-active" : ""}">
            <strong>Bajo</strong>
            <span>Temp <= 32 °C, viento <= 40 km/h y lluvia <= 20 mm.</span>
          </div>
          <div class="cold-threshold-item ${isMedio ? "is-active" : ""}">
            <strong>Medio</strong>
            <span>Si temp > 32 °C o viento > 40 km/h o lluvia > 20 mm.</span>
          </div>
          <div class="cold-threshold-item ${isAlto ? "is-active" : ""}">
            <strong>Alto</strong>
            <span>Si temp > 38 °C o viento > 55 km/h o lluvia > 40 mm.</span>
          </div>
        </div>
      </div>
    </div>
  `;
}

function formatRiskLevelLabel(level) {
  const normalized = normalizeRiskLevel(level);
  if (normalized === "critico") return "CRÍTICO";
  if (normalized === "alto") return "ALTO";
  if (normalized === "medio") return "MEDIO";
  return "BAJO";
}

function calcTotalRiskScore({ cadenaFria, logistico, regulatorio, meteorologico }) {
  const frioScore = riskLevelScore(cadenaFria);
  const logScore = riskLevelScore(logistico);
  const regScore = riskLevelScore(regulatorio);
  const metScore = riskLevelScore(meteorologico);
  const weighted = frioScore * 0.4 + logScore * 0.3 + regScore * 0.2 + metScore * 0.1;

  let totalLevel = "bajo";
  if (weighted >= 3.5) totalLevel = "critico";
  else if (weighted >= 2.5) totalLevel = "alto";
  else if (weighted >= 1.5) totalLevel = "medio";

  const hasCriticalComponent = [frioScore, logScore, regScore, metScore].some((value) => value >= 4);
  if (hasCriticalComponent && riskLevelScore(totalLevel) < riskLevelScore("alto")) {
    totalLevel = "alto";
  }

  return {
    score: weighted,
    level: totalLevel,
  };
}

function totalRiskAction(level) {
  const normalized = normalizeRiskLevel(level);
  if (normalized === "critico") return "Monitoreo continuo requerido antes del próximo cruce";
  if (normalized === "alto") return "Reforzar documentación térmica este mes";
  if (normalized === "medio") return "Monitoreo estándar recomendado";
  return "Operación estable";
}

function toggleRiesgoTotalThresholds(button) {
  if (!button) return;
  const panel = button.nextElementSibling;
  if (!panel) return;
  const isOpen = !panel.hidden;
  panel.hidden = isOpen;
  button.textContent = isOpen ? "Ver umbrales aplicados" : "Ocultar umbrales";
  button.setAttribute("aria-expanded", isOpen ? "false" : "true");
}

function renderRiesgoTotalBanner(totalRisk) {
  const banner = document.getElementById("riesgoTotalBanner");
  const text = document.getElementById("riesgoTotalText");
  const action = document.getElementById("riesgoTotalAction");
  const thresholdPanel = document.getElementById("riesgoTotalThresholdPanel");
  if (!banner || !text || !action) return;

  const level = normalizeRiskLevel(totalRisk?.level || "medio");
  const score = Number.isFinite(totalRisk?.score) ? totalRisk.score : 2;
  const label = formatRiskLevelLabel(level);

  banner.classList.remove("riesgo-total-critical", "riesgo-total-alto", "riesgo-total-medio", "riesgo-total-bajo");
  if (level === "critico") banner.classList.add("riesgo-total-critical");
  else if (level === "alto") banner.classList.add("riesgo-total-alto");
  else if (level === "medio") banner.classList.add("riesgo-total-medio");
  else banner.classList.add("riesgo-total-bajo");

  text.textContent = `Riesgo Total: ${label} (${score.toFixed(2)} / 4.0)`;
  action.textContent = totalRiskAction(level);

  if (thresholdPanel) {
    const isBajo = level === "bajo";
    const isMedio = level === "medio";
    const isAlto = level === "alto";
    const isCritico = level === "critico";
    thresholdPanel.innerHTML = `
      <div class="cold-thresholds-head">
        Fórmula aplicada: Cadena Fría (40%) + Logístico (30%) + Regulatorio (20%) + Meteorológico (10%).
      </div>
      <div class="cold-thresholds-list">
        <div class="cold-threshold-item ${isBajo ? "is-active" : ""}">
          <strong>Bajo</strong>
          <span>Score < 1.50</span>
        </div>
        <div class="cold-threshold-item ${isMedio ? "is-active" : ""}">
          <strong>Medio</strong>
          <span>Score 1.50 - 2.49</span>
        </div>
        <div class="cold-threshold-item ${isAlto ? "is-active" : ""}">
          <strong>Alto</strong>
          <span>Score 2.50 - 3.49</span>
        </div>
        <div class="cold-threshold-item ${isCritico ? "is-active" : ""}">
          <strong>Crítico</strong>
          <span>Score >= 3.50</span>
        </div>
      </div>
      <div class="cold-thresholds-applied">
        Regla adicional: si cualquier componente es crítico, el nivel total mínimo se eleva a alto.
      </div>
    `;
  }
}

function compactSentence(text) {
  const value = compactRouteText(text || "");
  if (!value || normalizeGeoKey(value) === normalizeGeoKey(NO_INFO)) return "";
  return value.replace(/\s+/g, " ").trim();
}

function getAduanaDisplayName(cruceInfo) {
  const direct = compactSentence(cruceInfo?.aduanaMx || "");
  if (direct) {
    const primary = direct.split("/")[0].trim();
    return primary || direct;
  }

  const nombre = compactSentence(cruceInfo?.nombre || "");
  if (!nombre) return "aduana no identificada";

  const mxSide = nombre.split("-")[0].trim();
  const cleaned = compactSentence(mxSide.replace(/\([^)]*\)/g, "").trim());
  return cleaned || mxSide || nombre;
}

function formatFdaHoursLabel(rawValue) {
  const value = compactSentence(rawValue || "");
  if (!value || !hasMeaningfulRoute(value)) return "No disponible";
  return value
    .replace(/\bhrs?\b\.?/gi, "")
    .replace(/\bhoras?\b\.?/gi, "")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function getDaysInMonthForDisplay(month, year = new Date().getFullYear()) {
  const numericMonth = Number(month);
  if (!Number.isFinite(numericMonth) || numericMonth < 1 || numericMonth > 12) return 31;
  return new Date(Number(year) || new Date().getFullYear(), numericMonth, 0).getDate();
}

function getColdProxyAnnualizedDays(proxy, month) {
  const yearsCount =
    Array.isArray(proxy?.years) && proxy.years.length
      ? proxy.years.length
      : COLD_PROXY_YEARS;
  const safeYears = Math.max(1, Number(yearsCount) || 1);
  const daysInMonth = getDaysInMonthForDisplay(month);
  const avgExtremeDays = Number.isFinite(proxy?.extremeDays) ? proxy.extremeDays / safeYears : null;
  return {
    yearsCount: safeYears,
    daysInMonth,
    avgExtremeDays,
  };
}

function buildRiesgoPitchText({ empresa, cruceInfo }) {
  const frecuenciaRaw = compactSentence(empresa?.frecuenciaSensor || "");
  const frecuencia = hasMeaningfulRoute(frecuenciaRaw) ? frecuenciaRaw : "5 minutos";
  const aduana = getAduanaDisplayName(cruceInfo);
  const tiempoFda = formatFdaHoursLabel(cruceInfo?.tiempoFda || "");

  return (
    `El sensor CLCircular registra temperatura cada ${frecuencia} durante las ${tiempoFda} hrs de inspección en ${aduana} ` +
    "y detecta automáticamente si abren las puertas del contenedor. " +
    "Al llegar a destino, genera un certificado descargable con la curva térmica completa del viaje — " +
    "el documento que tu importador en EE.UU. necesita para aprobar el embarque y cumplir FSMA 204."
  );
}

function calcColdChainRisk(clima) {
  const t = Math.max(...(clima.daily.temperature_2m_max || [0]));
  if (t > 35) return { level: "alto", label: "Alto" };
  if (t >= 30) return { level: "medio", label: "Medio" };
  return { level: "bajo", label: "Bajo" };
}

function syncRiesgoMesOperacionLabel() {
  const month = new Date().getMonth() + 1;
  const monthEl = document.getElementById("riesgoMesOperacionValue");
  if (!monthEl) return;
  monthEl.textContent = getMesOperacionLabel(month);
}

function getSelectedRiesgoMes() {
  return new Date().getMonth() + 1;
}

function getMesOperacionLabel(month) {
  const idx = Number(month) - 1;
  if (idx < 0 || idx >= RIESGO_MESES.length) return "mes no definido";
  return titleCase(RIESGO_MESES[idx]);
}

function higherRiskLevel(a, b) {
  const left = riskLevelScore(a);
  const right = riskLevelScore(b);
  return left >= right ? normalizeRiskLevel(a) : normalizeRiskLevel(b);
}

function normalizeRiskLevel(level) {
  const normalized = normalizeGeoKey(level);
  if (normalized.includes("CRITICO") || normalized.includes("CRITICA")) return "critico";
  if (normalized.includes("ALTO")) return "alto";
  if (normalized.includes("MEDIO") || normalized.includes("MODERADO")) return "medio";
  if (normalized.includes("BAJO")) return "bajo";
  return "medio";
}

function riskLevelScore(level) {
  const normalized = normalizeRiskLevel(level);
  if (normalized === "critico") return 4;
  if (normalized === "alto") return 3;
  if (normalized === "medio") return 2;
  return 1;
}

function parseEmpresaLogRiskCsv(riskText = "") {
  const raw = compactRouteText(riskText);
  if (!hasMeaningfulRoute(raw)) return null;

  const normalized = normalizeGeoKey(raw);
  let level = "medio";
  if (normalized.includes("CRITICO") || normalized.includes("CRITICA")) {
    level = "critico";
  } else if (normalized.includes("ALTO")) {
    level = "alto";
  } else if (normalized.includes("MODERADO") || normalized.includes("MEDIO")) {
    level = "medio";
  } else if (normalized.includes("BAJO")) {
    level = "bajo";
  }

  const reason = raw.includes(":") ? raw.split(":").slice(1).join(":").trim() : raw;
  const reasonSanitized = reason
    .replace(/Trayectos muy largos\s*\([^)]*Sinaloa[^)]*Arizona[^)]*\)\.?\s*/gi, "")
    .replace(/\s{2,}/g, " ")
    .trim();
  const detailByLevel = {
    critico: "Riesgo crítico; activar plan de contingencia antes del despacho.",
    alto: "Riesgo alto; reforzar control operativo antes del despacho.",
    medio: "Monitoreo reforzado; considerar margen adicional en despacho.",
    bajo: "Operación estable; monitoreo preventivo recomendado.",
  };

  return {
    level,
    label: titleCase(level),
    detail: detailByLevel[level] || detailByLevel.medio,
    mainReason: reasonSanitized || "",
    why: reasonSanitized ? `Riesgo logístico reportado por empresa: ${reasonSanitized}` : "",
  };
}

function mergeLogRiskSources(realtimeRisk, csvRisk) {
  if (!csvRisk) return realtimeRisk;
  if (!realtimeRisk) return csvRisk;

  const finalLevel = higherRiskLevel(realtimeRisk.level, csvRisk.level);
  const csvDominates = riskLevelScore(csvRisk.level) >= riskLevelScore(realtimeRisk.level);

  if (csvDominates) {
    return {
      level: finalLevel,
      label: titleCase(finalLevel),
      detail: csvRisk.detail,
      mainReason: csvRisk.mainReason || "",
      why: csvRisk.why,
    };
  }

  const mergedWhy = [realtimeRisk.why, csvRisk.why].filter(Boolean).join(" ");
  return {
    level: finalLevel,
    label: titleCase(finalLevel),
    detail: realtimeRisk.detail,
    mainReason: csvRisk.mainReason || realtimeRisk.mainReason || "",
    why: mergedWhy,
  };
}

function findCruceByEmpresa(empresa, rutaTerrestreText = "") {
  const cruces = (state.infraCruces || []).filter((item) => item && item.sourceType === "terrestre");
  if (!cruces.length) return null;

  const haystack = normalizeGeoKey(
    compactRouteText(
      `${empresa?.cruceFronterizo || ""} | ${empresa?.rutaTerrestre || ""} | ${rutaTerrestreText || ""} | ${
        empresa?.destinoFinal || empresa?.alcance || ""
      }`,
    ),
  );
  if (!haystack) return null;

  let best = null;
  let bestScore = 0;

  cruces.forEach((cruce) => {
    const keys = [
      cruce.nombre,
      cruce.aduanaMx,
      cruce.ciudadMx,
      cruce.ciudadUsa,
      cruce.estadoMx,
      cruce.estadoUsa,
      cruce.aduanaUsa,
    ]
      .map((value) => normalizeGeoKey(value))
      .filter(Boolean);

    let score = 0;
    keys.forEach((key) => {
      if (haystack.includes(key)) score += key.length >= 8 ? 4 : 2;
      const firstToken = key.split(" ")[0];
      if (firstToken && firstToken.length >= 5 && haystack.includes(firstToken)) score += 1;
    });

    if (score > bestScore) {
      bestScore = score;
      best = cruce;
    }
  });

  return bestScore > 0 ? best : null;
}

function findCruceByRutaReferencia(reference = "") {
  const cruces = (state.infraCruces || []).filter((item) => item && item.sourceType === "terrestre");
  if (!cruces.length) return null;
  const haystack = normalizeGeoKey(compactRouteText(reference));
  if (!haystack) return null;

  let best = null;
  let bestScore = 0;

  cruces.forEach((cruce) => {
    const keys = [cruce.nombre, cruce.aduanaMx, cruce.ciudadMx, cruce.ciudadUsa, cruce.estadoMx, cruce.estadoUsa]
      .map((value) => normalizeGeoKey(value))
      .filter(Boolean);

    let score = 0;
    keys.forEach((key) => {
      if (haystack.includes(key)) score += key.length >= 8 ? 6 : 3;
      const firstToken = key.split(" ")[0];
      if (firstToken && firstToken.length >= 5 && haystack.includes(firstToken)) score += 1;
    });

    if (score > bestScore) {
      bestScore = score;
      best = cruce;
    }
  });

  return bestScore > 0 ? best : null;
}

function setAduanaRiskPill(el, text, infraRiskLevel = "MODERADO") {
  if (!el) return;
  el.textContent = text;
  el.classList.remove("risk-aduana-critical", "risk-aduana-high", "risk-aduana-moderate");
  const normalized = normalizeInfraRiskLevel(infraRiskLevel);
  if (normalized === "CRITICO") {
    el.classList.add("risk-aduana-critical");
    return;
  }
  if (normalized === "ALTO") {
    el.classList.add("risk-aduana-high");
    return;
  }
  el.classList.add("risk-aduana-moderate");
}

function buildAduanaRiskCard(cruceInfo) {
  if (!cruceInfo) {
    return {
      level: "MODERADO",
      label: "No disponible",
      text: "No se detectó una aduana específica para esta empresa con los campos actuales.",
    };
  }

  const riskLevel = normalizeInfraRiskLevel(cruceInfo.riesgoLevel);
  const ftlMes = Number.isFinite(cruceInfo.ftlMensual)
    ? `${Math.round(cruceInfo.ftlMensual).toLocaleString("es-MX")} FTL/mes`
    : "No disponible";
  const tiempoFda = cruceInfo.tiempoFda || NO_INFO;
  const razon = compactRouteText(cruceInfo.riesgoReason || "");
  const hasRazon = razon && normalizeGeoKey(razon) !== normalizeGeoKey(NO_INFO);

  return {
    level: riskLevel,
    label: infraRiskLabel(riskLevel),
    text: `
      <span class="riesgo-aduana-line">
        <strong>${escapeHtml(cruceInfo.nombre || "Aduana no identificada")}</strong><br/>
        FTL del mes: ${escapeHtml(ftlMes)}<br/>
        Tiempo inspección FDA: ${escapeHtml(tiempoFda)}
        ${hasRazon ? `<br/>Razón del riesgo: ${escapeHtml(razon)}` : ""}
      </span>
    `,
  };
}

function limitWords(text, maxWords = 12) {
  const words = String(text || "")
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .filter(Boolean);
  if (!words.length) return "No disponible";
  if (words.length <= maxWords) return words.join(" ");
  return `${words.slice(0, maxWords).join(" ")}...`;
}

function normalizeCtpatYesNo(rawValue) {
  const normalized = normalizeGeoKey(rawValue);
  if (!normalized || normalized === normalizeGeoKey(NO_INFO)) return "No";
  if (/(SI|ACTIVO|TRUE|PARCIAL|PARTIAL)/.test(normalized)) return "Sí";
  if (/(NO|FALSE|INACTIVO)/.test(normalized)) return "No";
  return "No";
}

function buildRiskFieldRows(fields = []) {
  const items = fields
    .map((field) => {
      const label = escapeHtml(String(field?.label || "").trim());
      const valueText = compactRouteText(String(field?.value || ""));
      const value = escapeHtml(valueText || "No disponible");
      const valueClass = field?.compact ? "risk-field-value risk-field-value-short" : "risk-field-value";
      return `
        <div class="risk-field-item">
          <span class="risk-field-label">${label}</span>
          <span class="${valueClass}">${value}</span>
        </div>
      `;
    })
    .join("");

  return `<div class="risk-field-list">${items}</div>`;
}

function buildLogisticRiskFields({ logRiskCard, cruceInfo }) {
  const aduanaCruce = compactRouteText(cruceInfo?.nombre || "");
  const tiempoFda = compactRouteText(cruceInfo?.tiempoFda || "");
  const cTpat = normalizeCtpatYesNo(cruceInfo?.cTpatActivo || "");
  const riesgoRaw = compactRouteText(
    logRiskCard?.mainReason ||
      cruceInfo?.riesgoReason ||
      String(logRiskCard?.why || "").replace(/^Riesgo logístico reportado por empresa:\s*/i, "") ||
      logRiskCard?.detail ||
      "",
  );
  const riesgoPrincipal = limitWords(riesgoRaw, 12);

  return buildRiskFieldRows([
    { label: "Aduana de cruce", value: aduanaCruce || "No disponible" },
    { label: "Tiempo FDA", value: tiempoFda || "No disponible" },
    { label: "C-TPAT", value: cTpat },
    { label: "Riesgo principal", value: riesgoPrincipal, compact: true },
  ]);
}

function buildRegulatoryRiskFields(regRisk) {
  return buildRiskFieldRows([
    { label: "FSMA 204", value: regRisk?.fsmaRequired || "No requerido" },
    { label: "Rechazo FDA est.", value: regRisk?.rejectionEstimate || "No disponible" },
    { label: "Documentación temp.", value: regRisk?.tempDocumentation || "Estado desconocido — pendiente de verificación" },
    { label: "Acción recomendada", value: limitWords(regRisk?.recommendedAction || "", 12), compact: true },
  ]);
}

function estimateFdaRechazoByCruce(cruceInfo, normalizedRoute, hasUsSignals) {
  if (!hasUsSignals) return "No aplica";
  const haystack = normalizeGeoKey(
    `${cruceInfo?.nombre || ""} ${cruceInfo?.aduanaMx || ""} ${cruceInfo?.ciudadUsa || ""} ${normalizedRoute || ""}`,
  );
  if (!haystack) return "1-3%";
  if (haystack.includes("NOGALES")) return "3-5%";
  if (haystack.includes("NUEVO LAREDO") || haystack.includes("LAREDO")) return "2-4%";
  if (haystack.includes("OTAY") || haystack.includes("SAN YSIDRO") || haystack.includes("TIJUANA")) return "2-4%";
  if (haystack.includes("REYNOSA") || haystack.includes("PHARR") || haystack.includes("MCALLEN")) return "2-3%";
  if (haystack.includes("JUAREZ") || haystack.includes("EL PASO")) return "1-3%";
  if (haystack.includes("MATAMOROS") || haystack.includes("BROWNSVILLE")) return "2-3%";
  return "1-3%";
}

function resolveTempDocumentationStatus(hasTempRule, cTpatActivo) {
  if (!hasTempRule) return "Estado desconocido — pendiente de verificación";
  if (cTpatActivo) return "Automatizada";
  return "Manual";
}

function resolveRegulatoryAction(level, hasUsSignals) {
  if (!hasUsSignals) return "Mantener trazabilidad térmica y revisar requisitos antes de exportar.";
  if (level === "alto") return "Automatizar evidencia térmica y checklist FSMA antes del cruce.";
  if (level === "medio") return "Estandarizar bitácora térmica digital y validar documentos FSMA.";
  return "Mantener evidencia térmica digital y auditoría documental mensual.";
}

function calcRegulatoryRisk(empresa, logRiskLevel = "medio", cruceInfo = null) {
  const empresaText = compactRouteText(empresa?.empresa || "");
  const empresaNormalized = normalizeGeoKey(empresaText);
  const isGrupoPinsa = empresaNormalized.includes("GRUPO PINSA");
  const routeText = compactRouteText(
    `${empresa?.destinoFinal || ""} ${empresa?.alcance || ""} ${empresa?.cruceFronterizo || ""} ${
      empresa?.rutaTerrestre || ""
    } ${cruceInfo?.nombre || ""}`,
  );
  const normalizedRoute = normalizeGeoKey(routeText);
  const tempText = compactRouteText(empresa?.tempRequerida || "");
  const tempTextNorm = normalizeTempRiskText(tempText);
  const threshold = getCriticalThreshold(tempText);

  const hasUsSignals = [
    "EE UU",
    "ESTADOS UNIDOS",
    "USA",
    "TEXAS",
    "ARIZONA",
    "CALIFORNIA",
    "NOGALES",
    "LAREDO",
    "OTAY",
    "SAN YSIDRO",
    "PHARR",
    "MCALLEN",
    "BROWNSVILLE",
    "EL PASO",
    "LOS ANGELES",
    "SAN FRANCISCO",
  ].some((token) => normalizedRoute.includes(token));

  const hasTempRule =
    !!tempTextNorm &&
    (hasTempMarker(tempTextNorm, "4") ||
      hasTempMarker(tempTextNorm, "7") ||
      hasTempMarker(tempTextNorm, "-18") ||
      hasTempMarker(tempTextNorm, "-60") ||
      /(CONGELADO|FRESCO|REFRIGERADO|PASTEURIZADO|VIVO|TEMP\.?\s*AMBIENTE)/i.test(tempTextNorm));

  const cTpatRaw = normalizeGeoKey(cruceInfo?.cTpatActivo || "");
  const cTpatActivo =
    cTpatRaw.includes("SI") ||
    cTpatRaw.includes("ACTIVO") ||
    cTpatRaw.includes("TRUE") ||
    isGrupoPinsa;
  const isNogalesRoute = normalizedRoute.includes("NOGALES");
  const hasUsRetailSignal =
    normalizeGeoKey(`${empresa?.destinoFinal || ""} ${empresa?.alcance || ""}`).includes("RETAIL") ||
    normalizeGeoKey(`${empresa?.destinoFinal || ""} ${empresa?.alcance || ""}`).includes("FOOD SERVICE");
  let level = "medio";
  let detail = "Cumplimiento continuo requerido para exportar a EE.UU.";
  let why = "FSMA 204 exige trazabilidad y documentación de temperatura en tránsito.";

  if (isGrupoPinsa) {
    level = cTpatActivo ? "medio" : "alto";
    detail = "Cumplimiento regulatorio prioritario para exportación de Grupo Pinsa a EE.UU.";
    why = `C-TPAT: ${cTpatActivo ? "activo (reduce inspecciones ~80%)" : "no confirmado en cruce"} | FSMA 204: exporta a retailers de EE.UU.; compliance requerido | Tasa rechazo FDA en Nogales: ~3-5%.`;
  } else if (!hasUsSignals) {
    level = "bajo";
    detail = "Operación sin cruce activo a EE.UU.; cumplimiento base nacional.";
    why = "Mantener evidencia sanitaria y trazabilidad por lote para auditoría interna.";
  } else if (!hasTempRule) {
    level = "alto";
    detail = "Exportación a EE.UU. sin regla térmica explícita en ficha operativa.";
    why = "FSMA 204 requiere trazabilidad y respaldo de control de temperatura en tránsito.";
  } else {
    const severeLog = riskLevelScore(logRiskLevel) >= riskLevelScore("alto");
    if (severeLog || threshold <= 28 || isNogalesRoute) {
      level = "alto";
      detail = "Exigencia regulatoria alta por producto sensible y operación de frontera.";
      why = isNogalesRoute
        ? "FSMA 204 activo y cruce Nogales con rechazo FDA estimado de ~3-5%; priorizar bitácora térmica por embarque."
        : "FSMA 204 activo; priorizar bitácora térmica y evidencia digital por embarque.";
    } else if (hasUsRetailSignal) {
      level = "medio";
      detail = "Exportación a canales retail de EE.UU. con exigencia documental continua.";
      why = "FSMA 204 requiere trazabilidad lote a lote y respaldo de temperatura en tránsito.";
    }
  }

  const fsmaRequired = hasUsSignals ? "Requerido" : "No requerido";
  const rejectionEstimate = estimateFdaRechazoByCruce(cruceInfo, normalizedRoute, hasUsSignals);
  const tempDocumentation = isGrupoPinsa
    ? "Estado desconocido — pendiente de verificación"
    : resolveTempDocumentationStatus(hasTempRule, cTpatActivo);
  const recommendedAction = limitWords(resolveRegulatoryAction(level, hasUsSignals), 12);

  return {
    level,
    label: titleCase(level),
    detail,
    why,
    fsmaRequired,
    rejectionEstimate,
    tempDocumentation,
    recommendedAction,
  };
}

function normalizeTempRiskText(text) {
  return String(text || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[−–—]/g, "-")
    .replace(/≤/g, "<=")
    .replace(/≥/g, ">=")
    .replace(/º/g, "°")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function hasTempMarker(text, value) {
  const escaped = String(value)
    .replace(/[−–—]/g, "-")
    .replace(/\s+/g, "")
    .replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = `(^|[^0-9-])(?:<=|<|=)?\\s*${escaped}\\s*°?\\s*c\\b`;
  return new RegExp(pattern, "i").test(String(text || ""));
}

function getHeatThreshold(tempRequerida) {
  const text = normalizeTempRiskText(tempRequerida);
  if (!text) return COLD_PROXY_DEFAULT_THRESHOLD_C;

  if (text.includes("vivo") || hasTempMarker(text, "7")) {
    return 25;
  }
  if (
    hasTempMarker(text, "4") ||
    text.includes("fresco") ||
    text.includes("pasteurizado")
  ) {
    return 28;
  }
  if (hasTempMarker(text, "-18") || text.includes("congelado")) {
    return 32;
  }
  return COLD_PROXY_DEFAULT_THRESHOLD_C;
}

function getCriticalThreshold(tempRequeridaString) {
  const text = normalizeTempRiskText(tempRequeridaString);
  if (!text) return COLD_PROXY_DEFAULT_THRESHOLD_C;

  const thresholds = [];
  if (text.includes("vivo") || hasTempMarker(text, "7")) thresholds.push(25);
  if (hasTempMarker(text, "4") || text.includes("fresco") || text.includes("pasteurizado")) thresholds.push(28);
  if (hasTempMarker(text, "-18") || text.includes("congelado")) thresholds.push(32);

  if (thresholds.length) return Math.min(...thresholds);
  return Math.min(getHeatThreshold(text), COLD_PROXY_DEFAULT_THRESHOLD_C);
}

function buildColdProxyFallback(states, month, tempRequerida = "") {
  const thresholdC = getCriticalThreshold(tempRequerida);
  return {
    level: "medio",
    label: "Medio",
    probability: null,
    extremeDays: 0,
    totalDays: 0,
    avgMaxTemp: null,
    thresholdC,
    states: states || [],
    years: getColdProxyYears(month),
  };
}

function toggleColdThresholds(button) {
  if (!button) return;
  const panel = button.nextElementSibling;
  if (!panel) return;
  const isOpen = !panel.hidden;
  panel.hidden = isOpen;
  button.textContent = isOpen ? "Ver umbrales aplicados" : "Ocultar umbrales";
  button.setAttribute("aria-expanded", isOpen ? "false" : "true");
}

function buildColdChainProxyText({ tempRequerida, maxTemp, proxy, mesOperacion }) {
  const tempLabel = escapeHtml(formatTempDisplayText(tempRequerida, { includeFrozenAlso: true }));
  const maxTempLabel = Number.isFinite(maxTemp) ? escapeHtml(formatNumber(maxTemp, "°C")) : "No disponible";
  const monthLabel = escapeHtml(getMesOperacionLabel(mesOperacion));
  const annualizedDays = getColdProxyAnnualizedDays(proxy, mesOperacion);
  const thresholdC =
    proxy && Number.isFinite(proxy.thresholdC) ? Number(proxy.thresholdC) : COLD_PROXY_DEFAULT_THRESHOLD_C;
  const is25 = thresholdC === 25;
  const is28 = thresholdC === 28;
  const is32 = thresholdC === 32;
  const probLabel =
    proxy && Number.isFinite(proxy.probability)
      ? escapeHtml(formatPercentUnsigned(proxy.probability * 100))
      : "No disponible";
  const avgMaxLabel =
    proxy && Number.isFinite(proxy.avgMaxTemp) ? escapeHtml(formatNumber(proxy.avgMaxTemp, "°C")) : "No disponible";
  const diasExtremosDetalle =
    proxy && proxy.totalDays > 0 && Number.isFinite(annualizedDays.avgExtremeDays)
      ? `~${Math.round(annualizedDays.avgExtremeDays).toLocaleString("es-MX")} / ${annualizedDays.daysInMonth.toLocaleString("es-MX")} días (prom. ${annualizedDays.yearsCount.toLocaleString("es-MX")} años)`
      : "No disponible";

  return `
    <div class="cold-proxy-list">
      <div class="cold-proxy-item">
        <span class="cold-proxy-label">Temp requerida</span>
        <span class="cold-proxy-value">${tempLabel}</span>
      </div>
      <div class="cold-proxy-item">
        <span class="cold-proxy-label">Actual máx (7 días)</span>
        <span class="cold-proxy-value">${maxTempLabel}</span>
      </div>
      <div class="cold-proxy-item">
        <span class="cold-proxy-label">Días extremos (&gt;= ${formatNumber(thresholdC, "°C")}) en ${monthLabel}</span>
        <span class="cold-proxy-value">${diasExtremosDetalle}</span>
      </div>
      <div class="cold-proxy-item">
        <span class="cold-proxy-label">Temp máxima promedio mensual</span>
        <span class="cold-proxy-value">${avgMaxLabel}</span>
      </div>
      <div class="cold-proxy-item">
        <span class="cold-proxy-label">Probabilidad de descomposición por calor</span>
        <span class="cold-proxy-value">${probLabel}</span>
      </div>
    </div>
    <div class="cold-thresholds-wrap">
      <button
        type="button"
        class="cold-threshold-toggle"
        aria-expanded="false"
        onclick="toggleColdThresholds(this)"
      >
        Ver umbrales aplicados
      </button>
      <div class="cold-thresholds-panel" hidden>
        <div class="cold-thresholds-head">
          Regla: se usa el umbral más bajo del producto más sensible.
        </div>
        <div class="cold-thresholds-list">
          <div class="cold-threshold-item ${is25 ? "is-active" : ""}">
            <strong>25 °C</strong>
            <span>vivo o referencia 7°C</span>
          </div>
          <div class="cold-threshold-item ${is28 ? "is-active" : ""}">
            <strong>28 °C</strong>
            <span>fresco, pasteurizado o referencia 4°C</span>
          </div>
          <div class="cold-threshold-item ${is32 ? "is-active" : ""}">
            <strong>32 °C</strong>
            <span>congelado o referencia -18°C</span>
          </div>
        </div>
        <div class="cold-thresholds-applied">
          Umbral crítico aplicado: <strong>${formatNumber(thresholdC, "°C")}</strong>
          <br/>
          <small>
            Limitación: el símbolo "~" en días extremos indica un promedio anual estimado
            (extremos acumulados de ${annualizedDays.yearsCount.toLocaleString("es-MX")} años dividido entre
            ${annualizedDays.yearsCount.toLocaleString("es-MX")}), redondeado sobre los días del mes.
          </small>
        </div>
      </div>
    </div>
  `;
}

function resolveTransitStateKeys(empresa, rutaTerrestreNombre, rutaTerrestreDetalle) {
  const states = new Set();
  const sources = [
    empresa?.ubicacion || "",
    empresa?.sede || "",
    rutaTerrestreNombre || "",
    rutaTerrestreDetalle || "",
    empresa?.cruceFronterizo || "",
  ];
  sources.forEach((source) => {
    extractStateKeysFromText(source).forEach((state) => states.add(state));
  });
  return Array.from(states).filter((state) => estadoCoords[state]);
}

function extractStateKeysFromText(text) {
  const normalized = normalizeGeoKey(text);
  if (!normalized) return [];
  const tokens = new Set(normalized.split(" ").filter(Boolean));
  const states = new Set();

  Object.keys(estadoCoords).forEach((estado) => {
    if (estado === "MEXICO") return;
    if (normalized.includes(estado)) states.add(estado);
  });

  if (normalized.includes("ESTADO DE MEXICO") || normalized.includes("EDO MEX")) {
    states.add("MEXICO");
  }

  Object.entries(transitStateHints).forEach(([hint, state]) => {
    if (normalized.includes(hint)) states.add(state);
  });

  const abbrevMap = {
    SON: "SONORA",
    BC: "BAJA CALIFORNIA",
    BCS: "BAJA CALIFORNIA SUR",
    CDMX: "CIUDAD DE MEXICO",
  };
  Object.entries(abbrevMap).forEach(([abbr, state]) => {
    if (tokens.has(abbr)) states.add(state);
  });

  return Array.from(states);
}

function getColdProxyYears(month) {
  const currentDate = new Date();
  const currentYear = currentDate.getFullYear();
  const currentMonth = currentDate.getMonth() + 1;
  const useCurrentYear = Number(month) < currentMonth;
  const startYear = useCurrentYear ? currentYear : currentYear - 1;
  return Array.from({ length: COLD_PROXY_YEARS }, (_, idx) => startYear - idx);
}

async function calcColdChainProxyRisk(transitStates, month, tempRequerida = "") {
  const states = Array.from(new Set((transitStates || []).filter((state) => estadoCoords[state]))).slice(0, 7);
  const thresholdC = getCriticalThreshold(tempRequerida);
  if (!states.length) return buildColdProxyFallback([], month, tempRequerida);

  const years = getColdProxyYears(month);
  const tasks = [];
  states.forEach((state) => {
    const coord = estadoCoords[state];
    years.forEach((year) => {
      tasks.push(
        fetchMonthlyArchiveTemps(coord.lat, coord.lng, year, month).then((temps) => ({ state, year, temps })),
      );
    });
  });

  const series = await Promise.all(tasks);
  let extremeDays = 0;
  let totalDays = 0;
  let tempSum = 0;
  let tempCount = 0;

  series.forEach((item) => {
    (item.temps || []).forEach((temp) => {
      if (!Number.isFinite(temp)) return;
      totalDays += 1;
      tempSum += temp;
      tempCount += 1;
      if (temp >= thresholdC) extremeDays += 1;
    });
  });

  const probability = totalDays > 0 ? extremeDays / totalDays : null;
  let level = "medio";
  if (probability !== null) {
    if (probability >= 0.35) level = "alto";
    else if (probability >= 0.2) level = "medio";
    else level = "bajo";
  }

  return {
    level,
    label: titleCase(level),
    probability,
    extremeDays,
    totalDays,
    avgMaxTemp: tempCount > 0 ? tempSum / tempCount : null,
    thresholdC,
    states,
    years,
  };
}

async function fetchMonthlyArchiveTemps(lat, lng, year, month) {
  const mm = String(month).padStart(2, "0");
  const lastDay = String(new Date(year, month, 0).getDate()).padStart(2, "0");
  const startDate = `${year}-${mm}-01`;
  const endDate = `${year}-${mm}-${lastDay}`;
  const cacheKey = `${lat}|${lng}|${startDate}|${endDate}`;

  if (coldProxyArchiveCache.has(cacheKey)) return coldProxyArchiveCache.get(cacheKey);

  const fetchPromise = (async () => {
    try {
      const url =
        `https://archive-api.open-meteo.com/v1/archive?latitude=${lat}&longitude=${lng}` +
        `&start_date=${startDate}&end_date=${endDate}&daily=temperature_2m_max&timezone=auto`;
      const response = await fetch(url);
      if (!response.ok) throw new Error("Open-Meteo archive no respondió.");
      const data = await response.json();
      return data?.daily?.temperature_2m_max || [];
    } catch (error) {
      coldProxyArchiveCache.delete(cacheKey);
      throw error;
    }
  })();

  coldProxyArchiveCache.set(cacheKey, fetchPromise);
  return fetchPromise;
}

function calcWaveRisk(marine) {
  const maxWave = Math.max(...(marine.hourly?.wave_height || [0]));
  if (maxWave > 2.5) return { level: "alto", label: "alto", maxWave };
  if (maxWave > 1.5) return { level: "medio", label: "medio", maxWave };
  return { level: "bajo", label: "bajo", maxWave };
}

async function calcLogRiskRealtime(meteoLevel, rutaTerrestre, rutaTerrestreRaw = "") {
  const realtimeSignals = await fetchTrafficSignals(rutaTerrestre, rutaTerrestreRaw);

  const incidentSignal = realtimeSignals.find((signal) => signal.provider === "HERE" || signal.provider === "TomTom");
  if (incidentSignal) {
    let level = incidentSignal.level;
    if (meteoLevel === "alto" && level === "bajo") level = "medio";
    if (meteoLevel === "alto" && level === "medio") level = "alto";

    if (level === "alto") {
      return {
        level,
        label: "Alto",
        detail: "Riesgo alto de retraso logístico; ajustar ventana de salida.",
        why: `Porque ${incidentSignal.provider} reporta ${incidentSignal.incidents} incidentes y ${incidentSignal.closures} cierres en la ruta filtrada.`,
      };
    }
    if (level === "medio") {
      return {
        level,
        label: "Medio",
        detail: "Monitoreo reforzado; considerar margen adicional en despacho.",
        why: `Porque ${incidentSignal.provider} reporta ${incidentSignal.incidents} incidentes activos en la ruta filtrada.`,
      };
    }
    return {
      level,
      label: "Bajo",
      detail: "Operación estable; monitoreo preventivo recomendado.",
      why: `Porque ${incidentSignal.provider} no reporta incidentes críticos en la ruta filtrada.`,
    };
  }

  const cbpSignal = realtimeSignals.find((signal) => signal.provider === "CBP");
  if (cbpSignal && Number.isFinite(cbpSignal.waitMinutes)) {
    let level = "bajo";
    if (cbpSignal.waitMinutes >= 120) level = "alto";
    else if (cbpSignal.waitMinutes >= 60) level = "medio";

    if (meteoLevel === "alto" && level === "bajo") level = "medio";
    if (meteoLevel === "alto" && level === "medio") level = "alto";

    if (level === "alto") {
      return {
        level,
        label: "Alto",
        detail: "Riesgo alto de retraso logístico; ajustar ventana de salida.",
        why: `Porque CBP BWT reporta ~${formatNumber(cbpSignal.waitMinutes, "min")} en el cruce seleccionado en tiempo real.`,
      };
    }
    if (level === "medio") {
      return {
        level,
        label: "Medio",
        detail: "Monitoreo reforzado; considerar margen adicional en despacho.",
        why: `Porque CBP BWT reporta ~${formatNumber(cbpSignal.waitMinutes, "min")} en tiempo real y puede afectar tiempos de tránsito.`,
      };
    }
    return {
      level,
      label: "Bajo",
      detail: "Operación estable; monitoreo preventivo recomendado.",
      why: `Porque CBP BWT reporta ~${formatNumber(cbpSignal.waitMinutes, "min")} en tiempo real para el cruce seleccionado.`,
    };
  }

  return calcLogRiskFallback(meteoLevel, rutaTerrestre);
}

function calcLogRiskFallback(meteoLevel, rutaTerrestre) {
  const borderHeavy = rutaTerrestre.includes("Nuevo Laredo") || rutaTerrestre.includes("Reynosa");
  if (meteoLevel === "alto" || borderHeavy) {
    return {
      level: "medio",
      label: "Medio",
      detail: "Monitoreo reforzado; considerar margen adicional en despacho.",
      why: "Se aplicó evaluación preventiva por variabilidad operativa del corredor.",
    };
  }
  return {
    level: "bajo",
    label: "Bajo",
    detail: "Operación estable; monitoreo preventivo recomendado.",
    why: "Se aplicó evaluación preventiva con condiciones operativas estables.",
  };
}

async function fetchTrafficSignals(rutaTerrestre, rutaTerrestreRaw) {
  const signals = [];
  const bbox = buildTrafficBbox(rutaTerrestre, rutaTerrestreRaw);

  for (const provider of TRAFFIC_PROVIDER_PRIORITY) {
    if (provider === "here") {
      const here = await fetchHereIncidentsSignal(bbox);
      if (here) signals.push(here);
      continue;
    }
    if (provider === "tomtom") {
      const tomtom = await fetchTomTomIncidentsSignal(bbox);
      if (tomtom) signals.push(tomtom);
      continue;
    }
    if (provider === "cbp") {
      const wait = await fetchCbpWaitByCrossing(rutaTerrestre);
      if (Number.isFinite(wait)) signals.push({ provider: "CBP", waitMinutes: wait });
    }
  }

  return signals;
}

function buildTrafficBbox(rutaTerrestre, rutaTerrestreRaw) {
  const points = [];
  const fromRoute = extractGeoWaypointsFromRoute(rutaTerrestreRaw, "terrestre");
  fromRoute.forEach((point) => points.push({ lat: point.lat, lng: point.lng }));

  const crossing = borderCrossingCoords[rutaTerrestre];
  if (crossing) points.push({ lat: crossing.lat, lng: crossing.lng });

  if (!points.length) return null;
  let minLat = Number.POSITIVE_INFINITY;
  let minLng = Number.POSITIVE_INFINITY;
  let maxLat = Number.NEGATIVE_INFINITY;
  let maxLng = Number.NEGATIVE_INFINITY;

  points.forEach((point) => {
    minLat = Math.min(minLat, point.lat);
    minLng = Math.min(minLng, point.lng);
    maxLat = Math.max(maxLat, point.lat);
    maxLng = Math.max(maxLng, point.lng);
  });

  const padLat = 0.45;
  const padLng = 0.45;
  return {
    minLat: minLat - padLat,
    minLng: minLng - padLng,
    maxLat: maxLat + padLat,
    maxLng: maxLng + padLng,
  };
}

async function fetchHereIncidentsSignal(bbox) {
  if (!bbox || !HERE_TRAFFIC_API_KEY) return null;
  const inParam = `bbox:${bbox.minLat},${bbox.minLng},${bbox.maxLat},${bbox.maxLng}`;
  const url =
    `https://data.traffic.hereapi.com/v7/incidents?in=${encodeURIComponent(inParam)}` +
    `&locationReferencing=none&lang=es-ES&apiKey=${encodeURIComponent(HERE_TRAFFIC_API_KEY)}`;

  try {
    const payload = await fetchJsonWithTimeout(url, RISK_API_TIMEOUT_MS, "HERE Traffic");
    const incidents = Array.isArray(payload?.results) ? payload.results : Array.isArray(payload?.incidents) ? payload.incidents : [];
    const parsed = scoreIncidents(incidents);
    return {
      provider: "HERE",
      incidents: parsed.incidents,
      closures: parsed.closures,
      level: parsed.level,
    };
  } catch (error) {
    return null;
  }
}

async function fetchTomTomIncidentsSignal(bbox) {
  if (!bbox || !TOMTOM_TRAFFIC_API_KEY) return null;
  const bboxParam = `${bbox.minLng},${bbox.minLat},${bbox.maxLng},${bbox.maxLat}`;
  const url =
    `https://api.tomtom.com/traffic/services/5/incidentDetails?bbox=${encodeURIComponent(bboxParam)}` +
    `&language=es-ES&key=${encodeURIComponent(TOMTOM_TRAFFIC_API_KEY)}`;

  try {
    const payload = await fetchJsonWithTimeout(url, RISK_API_TIMEOUT_MS, "TomTom Traffic");
    const incidents = Array.isArray(payload?.incidents)
      ? payload.incidents
      : Array.isArray(payload?.tm?.poi)
        ? payload.tm.poi
        : [];
    const parsed = scoreIncidents(incidents);
    return {
      provider: "TomTom",
      incidents: parsed.incidents,
      closures: parsed.closures,
      level: parsed.level,
    };
  } catch (error) {
    return null;
  }
}

function scoreIncidents(incidents) {
  const rows = Array.isArray(incidents) ? incidents : [];
  const normalizedRows = rows.map((row) => normalizeGeoKey(JSON.stringify(row)));
  const closures = normalizedRows.filter((text) =>
    /(CLOSE|CLOSURE|CLOSED|BLOCK|BLOQUE|CERRAD|CARRETERA CERRADA)/i.test(text),
  ).length;
  const highSeverity = normalizedRows.filter((text) => /(CRITICAL|MAJOR|SEVERE|HIGH|ALTO)/i.test(text)).length;

  let level = "bajo";
  if (closures > 0 || highSeverity >= 3) level = "alto";
  else if (rows.length >= 5 || highSeverity > 0) level = "medio";

  return {
    incidents: rows.length,
    closures,
    level,
  };
}

async function fetchCbpWaitByCrossing(rutaTerrestre) {
  const rows = await fetchCbpWaitRows();
  if (!rows.length) return NaN;

  const keys = cbpCrossingKeys[rutaTerrestre] || [rutaTerrestre];
  const normalizedKeys = keys.map((k) => normalizeGeoKey(k));

  const matched = rows.filter((row) =>
    normalizedKeys.some((key) => row.name.includes(key) || key.includes(row.name)),
  );
  if (!matched.length) return NaN;

  const waits = matched.map((row) => row.wait).filter((value) => Number.isFinite(value) && value >= 0 && value <= 600);
  if (!waits.length) return NaN;
  return Math.max(...waits);
}

async function fetchCbpWaitRows() {
  const now = Date.now();
  if (state.cbpWaitCacheRows.length && now - state.cbpWaitCacheTs < CBP_CACHE_TTL_MS) {
    return state.cbpWaitCacheRows;
  }

  for (const endpoint of CBP_ENDPOINTS) {
    try {
      const payload = await fetchJsonWithTimeout(endpoint, CBP_FETCH_TIMEOUT_MS, "CBP BWT");
      const rows = parseCbpWaitRows(payload);
      if (rows.length) {
        state.cbpWaitCacheRows = rows;
        state.cbpWaitCacheTs = now;
        return rows;
      }
    } catch (error) {
      // Keep trying alternative endpoints.
    }
  }

  return [];
}

function parseCbpWaitRows(payload) {
  const objects = flattenObjects(payload);
  const rows = [];
  objects.forEach((entry) => {
    const name = extractBorderName(entry);
    const wait = extractCommercialWaitMinutes(entry);
    if (!name || !Number.isFinite(wait)) return;
    rows.push({ name: normalizeGeoKey(name), wait });
  });
  return rows;
}

function flattenObjects(value, out = []) {
  if (Array.isArray(value)) {
    value.forEach((item) => flattenObjects(item, out));
    return out;
  }
  if (!value || typeof value !== "object") return out;
  out.push(value);
  Object.values(value).forEach((item) => flattenObjects(item, out));
  return out;
}

function extractBorderName(entry) {
  const keys = [
    "port_name",
    "portName",
    "crossing_name",
    "crossingName",
    "name",
    "location",
    "port",
    "poe",
    "city",
  ];
  for (const key of keys) {
    const value = entry?.[key];
    if (typeof value === "string" && value.trim()) return value.trim();
  }
  return "";
}

function extractCommercialWaitMinutes(entry) {
  const values = collectNumericByPath(entry);
  const commercial = values
    .filter(
      (item) =>
        /(commercial|truck|cargo)/i.test(item.path) &&
        /(wait|delay|minute|time|min)/i.test(item.path),
    )
    .map((item) => item.value)
    .filter((value) => Number.isFinite(value) && value >= 0 && value <= 600);
  if (commercial.length) return Math.max(...commercial);

  const waitLike = values
    .filter((item) => /(wait|delay|minute|time|min)/i.test(item.path))
    .map((item) => item.value)
    .filter((value) => Number.isFinite(value) && value >= 0 && value <= 600);
  if (waitLike.length) return Math.max(...waitLike);

  return NaN;
}

function collectNumericByPath(value, path = "", out = []) {
  if (Array.isArray(value)) {
    value.forEach((item, idx) => collectNumericByPath(item, `${path}[${idx}]`, out));
    return out;
  }
  if (value && typeof value === "object") {
    Object.entries(value).forEach(([key, item]) => {
      const nextPath = path ? `${path}.${key}` : key;
      collectNumericByPath(item, nextPath, out);
    });
    return out;
  }

  const raw = String(value ?? "").trim();
  if (!raw) return out;
  const num = Number(raw.replace(/[^0-9.-]/g, ""));
  if (Number.isFinite(num)) {
    out.push({ path: path.toLowerCase(), value: num });
  }
  return out;
}

function setRiskPill(el, text, level) {
  if (!el) return;
  el.textContent = text;
  const normalized = normalizeRiskLevel(level);
  el.classList.remove("risk-bajo", "risk-medio", "risk-alto", "risk-critico");
  el.classList.add(`risk-${normalized}`);
}

function ensureLibs() {
  if (typeof XLSX === "undefined") {
    throw new Error("No cargó la librería XLSX.");
  }
}

function initTabs() {
  const buttons = document.querySelectorAll(".tab-button");
  const tabs = {
    kpis: document.getElementById("tab-kpis"),
    propuesta: document.getElementById("tab-propuesta"),
    viabilidad: document.getElementById("tab-viabilidad"),
    empresas: document.getElementById("tab-empresas"),
    riesgos: document.getElementById("tab-riesgos"),
    clustering: document.getElementById("tab-clustering"),
  };

  buttons.forEach((btn) => {
    btn.addEventListener("click", async () => {
      const tab = btn.dataset.tab;
      buttons.forEach((b) => b.classList.toggle("is-active", b === btn));
      Object.entries(tabs).forEach(([key, el]) => {
        if (!el) return;
        el.classList.toggle("is-active", key === tab);
      });
      updateFooterSources(tab);
      if (tab === "kpis") {
        await loadCompetidoresData();
        renderCompetidoresKpi();
      }
      if (tab === "empresas" || tab === "riesgos" || tab === "propuesta" || tab === "clustering") {
        await refreshEmpresasDataFromCsv();
      }
      if (tab === "riesgos") {
        await loadInfraData();
        initInfraKpi();
        initRiesgos();
      }
      if (tab === "propuesta") {
        renderPropuestaTab();
      }
      if (tab === "viabilidad") {
        renderViabilidadTab();
      }
      if (tab === "clustering") {
        renderClustering();
      }
      requestAnimationFrame(() => {
        resizeAllCharts();
        if (tab === "kpis" && state.kpiEstadosMap) {
          state.kpiEstadosMap.invalidateSize();
        }
        if (tab === "riesgos" && state.infraCrucesMap) {
          state.infraCrucesMap.invalidateSize();
        }
        if (tab === "kpis" && state.competidoresMap) {
          state.competidoresMap.invalidateSize();
        }
        if (tab === "empresas" && state.empresasMap) {
          state.empresasMap.invalidateSize();
        }
        if (tab === "propuesta" && state.propuestaCoverageMap) {
          state.propuestaCoverageMap.invalidateSize();
        }
        if (tab === "riesgos" && state.riesgoRouteMap) {
          state.riesgoRouteMap.invalidateSize();
        }
        if (tab === "riesgos" && state.riesgoMarineMap) {
          state.riesgoMarineMap.invalidateSize();
        }
      });
    });
  });

  updateFooterSources("kpis");
}

function updateFooterSources(tab) {
  const footer = document.getElementById("appFooterSources");
  if (!footer) return;
  syncFooterGlossaryByTab(tab);

  if (tab === "clustering") {
    footer.textContent = "Fuentes: empresas_pescados_mariscos_mexico_V2.csv";
    return;
  }

  if (tab === "riesgos") {
    footer.innerHTML = `
      <span>Fuentes:</span>
      <a href="https://open-meteo.com/" target="_blank" rel="noopener noreferrer">Open-Meteo</a>
      <span>|</span>
      <a href="https://bwt.cbp.gov/api/bwt" target="_blank" rel="noopener noreferrer">CBP BWT</a>
      <span>|</span>
      <a href="https://project-osrm.org/" target="_blank" rel="noopener noreferrer">OSRM (distancia por carretera)</a>
      <span>|</span>
      <a href="https://www.nhc.noaa.gov/" target="_blank" rel="noopener noreferrer">NOAA/NHC</a>
      <span>|</span>
      <a href="https://smn.conagua.gob.mx/" target="_blank" rel="noopener noreferrer">SMN México</a>
    `;
    return;
  }

  if (tab === "propuesta") {
    footer.textContent =
      "Fuentes: CONAPESCA_BASE_UNIFICADA_2005_2024 | empresas_pescados_mariscos_mexico_V2.csv | Cruces Terrestres.csv";
    return;
  }

  footer.textContent = "Fuentes: CONAPESCA_BASE_UNIFICADA_2005_2024";
}

function syncFooterGlossaryByTab(tab) {
  const button = document.getElementById("footerGlossaryBtn");
  const panel = document.getElementById("footerGlossaryPanel");
  if (!button || !panel) return;

  if (tab === "clustering") {
    button.setAttribute("hidden", "");
    panel.setAttribute("hidden", "");
    button.setAttribute("aria-expanded", "false");
    button.textContent = "Glosario ▼";
    return;
  }

  button.removeAttribute("hidden");
  panel.setAttribute("hidden", "");
  button.setAttribute("aria-expanded", "false");
  button.textContent = "Glosario ▼";
}

function initFooterGlossary() {
  const button = document.getElementById("footerGlossaryBtn");
  const panel = document.getElementById("footerGlossaryPanel");
  if (!button || !panel) return;
  setFooterGlossaryState(false);
}

function toggleFooterGlossary() {
  const button = document.getElementById("footerGlossaryBtn");
  const panel = document.getElementById("footerGlossaryPanel");
  if (!button || !panel) return;
  const open = !panel.hasAttribute("hidden");
  setFooterGlossaryState(!open);
}

function setFooterGlossaryState(open) {
  const button = document.getElementById("footerGlossaryBtn");
  const panel = document.getElementById("footerGlossaryPanel");
  if (!button || !panel) return;
  if (open) {
    panel.removeAttribute("hidden");
    button.setAttribute("aria-expanded", "true");
    button.textContent = "Glosario ▲";
    return;
  }
  panel.setAttribute("hidden", "");
  button.setAttribute("aria-expanded", "false");
  button.textContent = "Glosario ▼";
}

function toggleRiesgoClCircularInfo() {
  const button = document.getElementById("riesgoClCircularBtn");
  const panel = document.getElementById("riesgoClCircularPanel");
  if (!button || !panel) return;
  const isOpen = !panel.hasAttribute("hidden");
  if (isOpen) {
    panel.setAttribute("hidden", "");
    button.setAttribute("aria-expanded", "false");
    button.textContent = "Qué hace CL Circular";
    return;
  }
  panel.removeAttribute("hidden");
  button.setAttribute("aria-expanded", "true");
  button.textContent = "Ocultar qué hace CL Circular";
}

function renderForecast7d(clima) {
  const grid = document.getElementById("riesgoPronosticoGrid");
  if (!grid) return;
  if (!clima?.daily?.time?.length) {
    grid.innerHTML = "";
    return;
  }

  const times = clima.daily.time || [];
  const tMax = clima.daily.temperature_2m_max || [];
  const tMin = clima.daily.temperature_2m_min || [];
  const lluvia = clima.daily.precipitation_sum || [];
  const viento = clima.daily.wind_speed_10m_max || [];
  const weather = clima.daily.weather_code || [];

  grid.innerHTML = times
    .slice(0, 7)
    .map((date, i) => {
      const d = new Date(`${date}T12:00:00`);
      const dayLabel = d.toLocaleDateString("es-MX", { weekday: "short", day: "2-digit", month: "2-digit" });
      const weatherInfo = weatherCodeToLabel(weather[i]);
      const weatherPrefix = weatherInfo.icon ? `<span class="forecast-icon">${weatherInfo.icon}</span> ` : "";
      return `
        <div class="forecast-item">
          <div class="forecast-day">${dayLabel}</div>
          <div>${weatherPrefix}${weatherInfo.label}</div>
          <div>Max: ${formatNumber(tMax[i] ?? 0, "°C")}</div>
          <div>Min: ${formatNumber(tMin[i] ?? 0, "°C")}</div>
          <div>Lluvia: ${formatNumber(lluvia[i] ?? 0, "mm")}</div>
          <div>Viento: ${formatNumber(viento[i] ?? 0, "km/h")}</div>
        </div>
      `;
    })
    .join("");
}

function weatherCodeToLabel(code) {
  const c = Number(code);
  if (c === 0) return { icon: "", label: "Despejado" };
  if ([1, 2].includes(c)) return { icon: "", label: "Parcial nublado" };
  if (c === 3) return { icon: "", label: "Nublado" };
  if ([45, 48].includes(c)) return { icon: "", label: "Neblina" };
  if ([51, 53, 55, 56, 57].includes(c)) return { icon: "", label: "Llovizna" };
  if ([61, 63, 65, 66, 67, 80, 81, 82].includes(c)) return { icon: "", label: "Lluvia" };
  if ([71, 73, 75, 77, 85, 86].includes(c)) return { icon: "", label: "Nieve" };
  if ([95, 96, 99].includes(c)) return { icon: "", label: "Tormenta" };
  return { icon: "", label: "Variable" };
}

function renderRiesgoRutaMap(empresa, rutaTerrestre, nivelRiesgo, rutaTerrestreRaw = "") {
  const mapEl = document.getElementById("riesgoRutaTerrestreMap");
  const distEl = document.getElementById("riesgoRutaTerrestreDist");
  const leyendaEl = document.getElementById("riesgoRutaLeyenda");
  if (!mapEl || typeof L === "undefined") return;

  if (state.riesgoRouteMap) {
    state.riesgoRouteMap.remove();
    state.riesgoRouteMap = null;
    state.riesgoRouteLayers = [];
  }

  const crossing =
    borderCrossingCoords[rutaTerrestre.nombre] || borderCrossingCoords["Nuevo Laredo (Tamaulipas) - Laredo (Texas)"];
  const usDest =
    usDestinationsByCrossing[rutaTerrestre.nombre] || usDestinationsByCrossing["Nuevo Laredo (Tamaulipas) - Laredo (Texas)"];
  const roadPivot = midpoint(empresa, crossing);
  const csvWaypoints = extractGeoWaypointsFromRoute(rutaTerrestreRaw, "terrestre");

  const map = L.map(mapEl).setView([empresa.lat, empresa.lng], 4.6);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);

  let segments = [];
  if (csvWaypoints.length >= 2) {
    const points = csvWaypoints;
    segments = points.slice(0, -1).map((point, idx) => {
      const next = points[idx + 1];
      const level = idx === points.length - 2 ? routeSegmentLevel("cruce", nivelRiesgo) : routeSegmentLevel("carretera", nivelRiesgo);
      return {
        label: `${point.label} -> ${next.label}`,
        points: [
          [point.lat, point.lng],
          [next.lat, next.lng],
        ],
        level,
      };
    });
  } else {
    segments = [
      {
        label: "Origen -> carretera",
        points: [
          [empresa.lat, empresa.lng],
          [roadPivot.lat, roadPivot.lng],
        ],
        level: routeSegmentLevel("origen", nivelRiesgo),
      },
      {
        label: "Carretera -> cruce",
        points: [
          [roadPivot.lat, roadPivot.lng],
          [crossing.lat, crossing.lng],
        ],
        level: routeSegmentLevel("carretera", nivelRiesgo),
      },
      {
        label: "Cruce -> destino EE.UU.",
        points: [
          [crossing.lat, crossing.lng],
          [usDest.lat, usDest.lng],
        ],
        level: routeSegmentLevel("cruce", nivelRiesgo),
      },
    ];
  }

  const totalDistanceKm = totalRouteDistanceKm(segments);
  const routeCoords = buildRouteCoordsFromSegments(segments);
  const distanceReqKey = routeCoordsCacheKey(routeCoords);
  if (distEl) {
    distEl.dataset.requestKey = distanceReqKey;
    distEl.textContent = "Distancia por carretera al destino: calculando...";
  }
  fetchRoadDistanceKmFromOsrm(routeCoords).then((roadDistanceKm) => {
    if (!distEl) return;
    if ((distEl.dataset.requestKey || "") !== distanceReqKey) return;
    if (Number.isFinite(roadDistanceKm) && roadDistanceKm > 0) {
      distEl.textContent = `Distancia por carretera al destino: ${formatDistanceKm(roadDistanceKm)} (OSRM)`;
      return;
    }
    distEl.textContent = `Distancia estimada al destino: ${formatDistanceKm(totalDistanceKm)}`;
  });

  if (leyendaEl) {
    leyendaEl.innerHTML = `
      <div class="leyenda-card">
        <span class="leyenda-dot alto"></span>
        <div>
          <strong>Alto</strong>
          <small>Retraso/impacto alto</small>
        </div>
      </div>
      <div class="leyenda-card">
        <span class="leyenda-dot medio"></span>
        <div>
          <strong>Medio</strong>
          <small>Monitoreo reforzado</small>
        </div>
      </div>
      <div class="leyenda-card">
        <span class="leyenda-dot bajo"></span>
        <div>
          <strong>Bajo</strong>
          <small>Operacion estable</small>
        </div>
      </div>
    `;
  }

  const allPoints = [];
  segments.forEach((seg) => {
    const color = riskColor(seg.level);
    const poly = L.polyline(seg.points, { color, weight: 5, opacity: 0.85 }).addTo(map);
    poly.bindTooltip(`${seg.label}: ${titleCase(seg.level)}`);
    state.riesgoRouteLayers.push(poly);
    allPoints.push(...seg.points);
  });

  if (csvWaypoints.length >= 2) {
    csvWaypoints.forEach((point) => {
      L.marker([point.lat, point.lng]).addTo(map).bindPopup(point.label);
    });
  } else {
    L.marker([empresa.lat, empresa.lng]).addTo(map).bindPopup(`Origen: ${empresa.empresa}`);
    L.marker([crossing.lat, crossing.lng]).addTo(map).bindPopup(`Cruce: ${rutaTerrestre.nombre}`);
    L.marker([usDest.lat, usDest.lng]).addTo(map).bindPopup(`Destino: ${usDest.nombre}`);
  }

  map.fitBounds(allPoints, { padding: [110, 90], maxZoom: 5 });
  state.riesgoRouteMap = map;
}

function renderRiesgoRutaMaritimaMap(empresa, rutaOceanica, waveLevel, rutaMaritimaRaw = "") {
  const mapEl = document.getElementById("riesgoRutaMaritimaMap");
  const distEl = document.getElementById("riesgoRutaMaritimaDist");
  if (!mapEl || typeof L === "undefined") return;

  if (state.riesgoMarineMap) {
    state.riesgoMarineMap.remove();
    state.riesgoMarineMap = null;
    state.riesgoMarineLayers = [];
  }

  const csvWaypoints = extractGeoWaypointsFromRoute(rutaMaritimaRaw, "maritima");
  const puerto = puertoCoords[rutaOceanica.nombre];
  const destinoUs =
    usMaritimeDestByPuerto[rutaOceanica.nombre] || { nombre: "Puerto en EE.UU.", lat: 29.7604, lng: -95.3698 };
  const pivotCosta = puerto ? midpoint(empresa, puerto) : null;

  const startCenter = csvWaypoints[0] || puerto || { lat: 23.5, lng: -102.5 };
  const map = L.map(mapEl).setView([startCenter.lat, startCenter.lng], 4.6);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);

  let segments = [];
  if (csvWaypoints.length >= 2) {
    segments.push({
      points: [
        [empresa.lat, empresa.lng],
        [csvWaypoints[0].lat, csvWaypoints[0].lng],
      ],
      color: "#1f77b4",
      label: "Acceso al puerto",
    });
    csvWaypoints.slice(0, -1).forEach((point, idx) => {
      const next = csvWaypoints[idx + 1];
      segments.push({
        points: [
          [point.lat, point.lng],
          [next.lat, next.lng],
        ],
        color: riskColor(waveLevel || "medio"),
        label: `${point.label} -> ${next.label}`,
      });
    });
  } else if (puerto && pivotCosta) {
    segments = [
      {
        points: [
          [empresa.lat, empresa.lng],
          [pivotCosta.lat, pivotCosta.lng],
          [puerto.lat, puerto.lng],
        ],
        color: "#1f77b4",
        label: "Acceso al puerto",
      },
      {
        points: [
          [puerto.lat, puerto.lng],
          [destinoUs.lat, destinoUs.lng],
        ],
        color: riskColor(waveLevel || "medio"),
        label: `Tramo marítimo (${titleCase(waveLevel || "medio")})`,
      },
    ];
  } else {
    L.marker([empresa.lat, empresa.lng]).addTo(map).bindPopup(`Origen: ${empresa.empresa}`);
    if (distEl) {
      distEl.textContent = "Distancia estimada al destino: no disponible";
    }
    state.riesgoMarineMap = map;
    return;
  }

  const totalDistanceKm = totalRouteDistanceKm(segments);
  if (distEl) {
    distEl.textContent = `Distancia estimada al destino: ${formatDistanceKm(totalDistanceKm)}`;
  }

  const bounds = [];
  segments.forEach((seg) => {
    const poly = L.polyline(seg.points, { color: seg.color, weight: 4, opacity: 0.85, dashArray: "8, 6" }).addTo(map);
    poly.bindTooltip(seg.label);
    state.riesgoMarineLayers.push(poly);
    bounds.push(...seg.points);
  });

  L.marker([empresa.lat, empresa.lng]).addTo(map).bindPopup(`Origen: ${empresa.empresa}`);
  if (csvWaypoints.length >= 2) {
    csvWaypoints.forEach((point) => {
      L.marker([point.lat, point.lng]).addTo(map).bindPopup(point.label);
    });
  } else {
    L.marker([puerto.lat, puerto.lng]).addTo(map).bindPopup(`Puerto: ${rutaOceanica.nombre}`);
    L.marker([destinoUs.lat, destinoUs.lng]).addTo(map).bindPopup(`Destino marítimo: ${destinoUs.nombre}`);
  }

  map.fitBounds(bounds, { padding: [110, 90], maxZoom: 5 });
  state.riesgoMarineMap = map;
}

function renderRiesgoRutaMaritimaNoAplica() {
  const mapEl = document.getElementById("riesgoRutaMaritimaMap");
  const distEl = document.getElementById("riesgoRutaMaritimaDist");
  if (!mapEl || typeof L === "undefined") return;

  if (state.riesgoMarineMap) {
    state.riesgoMarineMap.remove();
    state.riesgoMarineMap = null;
    state.riesgoMarineLayers = [];
  }

  const map = L.map(mapEl).setView([23.5, -102.5], 4.6);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);
  L.marker([23.5, -102.5]).addTo(map).bindPopup("Ruta marítima no aplica para esta empresa");
  if (distEl) {
    distEl.textContent = "Distancia estimada al destino: no aplica";
  }
  state.riesgoMarineMap = map;
}

function midpoint(a, b) {
  return { lat: (a.lat + b.lat) / 2, lng: (a.lng + b.lng) / 2 };
}

function buildRouteCoordsFromSegments(segments = []) {
  const coords = [];
  (segments || []).forEach((segment) => {
    (segment?.points || []).forEach((point) => {
      if (!Array.isArray(point) || point.length < 2) return;
      const lat = Number(point[0]);
      const lng = Number(point[1]);
      if (!Number.isFinite(lat) || !Number.isFinite(lng)) return;
      const prev = coords[coords.length - 1];
      if (!prev || Math.abs(prev.lat - lat) > 1e-6 || Math.abs(prev.lng - lng) > 1e-6) {
        coords.push({ lat, lng });
      }
    });
  });
  return coords;
}

function compressRouteCoords(coords = [], maxCoords = ROAD_DISTANCE_MAX_COORDS) {
  const points = (coords || []).filter((p) => Number.isFinite(Number(p?.lat)) && Number.isFinite(Number(p?.lng)));
  if (points.length <= maxCoords) return points;

  const sampled = [points[0]];
  const step = (points.length - 1) / (maxCoords - 1);
  for (let i = 1; i < maxCoords - 1; i += 1) {
    sampled.push(points[Math.round(i * step)]);
  }
  sampled.push(points[points.length - 1]);

  return sampled.filter((point, idx, arr) => {
    if (idx === 0) return true;
    const prev = arr[idx - 1];
    return Math.abs(point.lat - prev.lat) > 1e-6 || Math.abs(point.lng - prev.lng) > 1e-6;
  });
}

function routeCoordsCacheKey(coords = []) {
  const compact = compressRouteCoords(coords);
  return compact.map((point) => `${point.lat.toFixed(4)},${point.lng.toFixed(4)}`).join(";");
}

async function fetchRoadDistanceKmFromOsrm(coords = []) {
  const compact = compressRouteCoords(coords);
  if (compact.length < 2) return null;

  const cacheKey = routeCoordsCacheKey(compact);
  if (roadDistanceCache.has(cacheKey)) return roadDistanceCache.get(cacheKey);

  const coordString = compact.map((point) => `${point.lng},${point.lat}`).join(";");
  const url = `${OSRM_ROUTE_BASE_URL}/${coordString}?overview=false&alternatives=false&steps=false`;
  const controller = typeof AbortController !== "undefined" ? new AbortController() : null;
  const timeoutId = controller ? setTimeout(() => controller.abort(), ROAD_DISTANCE_TIMEOUT_MS) : null;

  try {
    const res = await fetch(url, { cache: "no-store", signal: controller ? controller.signal : undefined });
    if (!res.ok) return null;
    const payload = await res.json();
    const distanceMeters = Number(payload?.routes?.[0]?.distance);
    if (!Number.isFinite(distanceMeters) || distanceMeters <= 0) return null;
    const distanceKm = distanceMeters / 1000;
    roadDistanceCache.set(cacheKey, distanceKm);
    return distanceKm;
  } catch (error) {
    return null;
  } finally {
    if (timeoutId) clearTimeout(timeoutId);
  }
}

function haversineKm(from, to) {
  const lat1 = Number(from?.lat);
  const lng1 = Number(from?.lng);
  const lat2 = Number(to?.lat);
  const lng2 = Number(to?.lng);
  if (![lat1, lng1, lat2, lng2].every(Number.isFinite)) return 0;

  const toRad = (deg) => (deg * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return 6371 * c;
}

function polylineDistanceKm(points = []) {
  if (!Array.isArray(points) || points.length < 2) return 0;
  let total = 0;
  for (let i = 0; i < points.length - 1; i += 1) {
    const current = points[i];
    const next = points[i + 1];
    total += haversineKm({ lat: current[0], lng: current[1] }, { lat: next[0], lng: next[1] });
  }
  return total;
}

function totalRouteDistanceKm(segments = []) {
  if (!Array.isArray(segments) || !segments.length) return 0;
  return segments.reduce((acc, seg) => acc + polylineDistanceKm(seg?.points || []), 0);
}

function formatDistanceKm(valueKm) {
  const value = Number(valueKm);
  if (!Number.isFinite(value) || value <= 0) return "no disponible";
  return `${value.toLocaleString("es-MX", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  })} km`;
}

function routeSegmentLevel(segment, base) {
  if (segment === "cruce") {
    if (base === "alto") return "alto";
    return "medio";
  }
  if (segment === "carretera") return base;
  if (base === "alto") return "medio";
  return "bajo";
}

function riskColor(level) {
  if (level === "alto") return "#d62828";
  if (level === "medio") return "#f39c12";
  return "#046f31";
}

function initClustering() {
  // Tab intentionally left empty.
}

let pcaVisible = false;
let clusterFeaturesVisible = false;

function setClusterFeaturesVisibility() {
  const btn = document.getElementById("clusterFeaturesToggleBtn");
  const panel = document.getElementById("clusterFeaturesPanel");
  const visible = !!clusterFeaturesVisible;
  if (btn) {
    btn.classList.toggle("is-active", visible);
    btn.setAttribute("aria-expanded", visible ? "true" : "false");
  }
  if (panel) panel.hidden = !visible;
}

function initClusterFeaturesToggle() {
  const btn = document.getElementById("clusterFeaturesToggleBtn");
  if (!btn) return;
  if (btn.dataset.bound === "1") {
    setClusterFeaturesVisibility();
    return;
  }
  btn.dataset.bound = "1";
  btn.addEventListener("click", () => {
    clusterFeaturesVisible = !clusterFeaturesVisible;
    setClusterFeaturesVisibility();
  });
  setClusterFeaturesVisibility();
}

function setClusterPcaVisibility() {
  const btn = document.getElementById("clusterPcaToggleBtn");
  const pcaEl = document.getElementById("clusterPcaPlot");
  const scatterEl = document.getElementById("clusterScatterPlot");
  const visible = !!pcaVisible;
  if (btn) {
    btn.classList.toggle("is-active", visible);
    btn.setAttribute("aria-pressed", visible ? "true" : "false");
    btn.textContent = visible ? "Ver scatter plot" : "Proyección PCA 2D";
  }
  if (pcaEl) pcaEl.hidden = !visible;
  if (scatterEl) scatterEl.hidden = visible;
}

function initClusterPcaToggle() {
  const btn = document.getElementById("clusterPcaToggleBtn");
  if (!btn) return;
  if (btn.dataset.bound === "1") {
    setClusterPcaVisibility();
    return;
  }
  btn.dataset.bound = "1";
  btn.addEventListener("click", () => {
    pcaVisible = !pcaVisible;
    setClusterPcaVisibility();
    renderClustering();
  });
  setClusterPcaVisibility();
}

const CLUSTER_PRIORITY_NAMES = [
  "Grupo Pinsa",
  "Baja Shellfish Farms",
  "Baja Aqua-Farms",
  "Grupo Acuícola Mexicano (GAM)",
  "Pacifico Aquaculture",
];

const CLUSTER_METRIC_OVERRIDES = {
  "BAJA SHELLFISH FARMS": {
    ftlMes: 40,
    ventasMUsd: 22,
    riesgoLogistico: 4,
    cruceScore: 3,
    numCertificaciones: 3,
  },
};

const CLUSTER_SALES_BASELINE_MUSD = {
  "GRUPO PINSA": 125,
  "GAM": 40,
  "GRUPO ACUICOLA MEXICANO GAM": 40,
  "BAJA AQUA FARMS": 100,
  "PACIFICO AQUACULTURE": 50,
  "BAJA SHELLFISH FARMS": 22,
};

const CLUSTER_BASE_ROWS = [
  { empresa: "Grupo Pinsa", ftlLabel: "42.8", ftlValue: 42.8, aduana: "Nogales", riesgo: "MODERADO", quarter: "Q2 2026" },
  { empresa: "GAM", ftlLabel: "60.0", ftlValue: 60.0, aduana: "Nogales", riesgo: "ALTO", quarter: "Q2 2026" },
  {
    empresa: "Baja Aqua-Farms",
    ftlLabel: "88.3",
    ftlValue: 88.3,
    aduana: "Otay Mesa",
    riesgo: "CRÍTICO",
    quarter: "Q3 2026",
  },
  {
    empresa: "Pacífico Aquaculture",
    ftlLabel: "24.7",
    ftlValue: 24.7,
    aduana: "Otay Mesa",
    riesgo: "ALTO",
    quarter: "Q3 2026",
  },
  {
    empresa: "Baja Shellfish Farms",
    ftlLabel: "40.0",
    ftlValue: 40.0,
    aduana: "Otay Mesa",
    riesgo: "CRÍTICO",
    quarter: "Q4 2026",
  },
];

const ESTRATEGIA_PROSPECT_OVERRIDES = {
  "BAJA AQUA FARMS": { profile: "Ancla", quarter: "Q2 2026" },
  "GRUPO PINSA": { profile: "Estratégico", quarter: "Q3 2026" },
  "GAM": { profile: "Estratégico", quarter: "Q3 2026" },
  "GRUPO ACUICOLA MEXICANO GAM": { profile: "Estratégico", quarter: "Q3 2026" },
  "PACIFICO AQUACULTURE": { profile: "Estratégico", quarter: "Q4 2026" },
  "BAJA SHELLFISH FARMS": { profile: "Estratégico", quarter: "Q4 2026" },
};

function normalizeCompanyKey(name) {
  return normalizeGeoKey(name);
}

function getEstrategiaProspectOverride(name = "") {
  const key = normalizeCompanyKey(name);
  if (!key) return null;
  if (ESTRATEGIA_PROSPECT_OVERRIDES[key]) return ESTRATEGIA_PROSPECT_OVERRIDES[key];
  const matchKey = Object.keys(ESTRATEGIA_PROSPECT_OVERRIDES).find(
    (overrideKey) => key.includes(overrideKey) || overrideKey.includes(key),
  );
  return matchKey ? ESTRATEGIA_PROSPECT_OVERRIDES[matchKey] : null;
}

function selectPriorityClusterCompanies(rows = []) {
  const byKey = new Map(
    (rows || []).map((empresa) => [normalizeCompanyKey(empresa?.empresa || ""), empresa]).filter(([k]) => !!k),
  );

  const selected = [];
  CLUSTER_PRIORITY_NAMES.forEach((name) => {
    const targetKey = normalizeCompanyKey(name);
    const exact = byKey.get(targetKey);
    if (exact) {
      selected.push(exact);
      return;
    }
    const fuzzy = (rows || []).find((empresa) => normalizeCompanyKey(empresa?.empresa || "").includes(targetKey));
    if (fuzzy) selected.push(fuzzy);
  });

  if (selected.length >= 3) return selected;
  return (rows || []).slice(0, 5);
}

function parseRangeAverage(text) {
  const raw = compactRouteText(text);
  if (!raw) return NaN;
  const rangeMatch = raw.match(/(\d+(?:[.,]\d+)?)\s*(?:a|-|–|—)\s*(\d+(?:[.,]\d+)?)/i);
  if (rangeMatch) {
    const min = Number(rangeMatch[1].replace(/,/g, ""));
    const max = Number(rangeMatch[2].replace(/,/g, ""));
    if (Number.isFinite(min) && Number.isFinite(max)) return (min + max) / 2;
  }
  const singleMatch = raw.match(/(\d+(?:[.,]\d+)?)/);
  if (!singleMatch) return NaN;
  const single = Number(singleMatch[1].replace(/,/g, ""));
  return Number.isFinite(single) ? single : NaN;
}

function parseSalesMUsd(text) {
  const raw = compactRouteText(text);
  if (!raw) return NaN;
  const matches = Array.from(raw.matchAll(/(\d+(?:[.,]\d+)?)\s*M\b/gi))
    .map((m) => Number(m[1].replace(/,/g, "")))
    .filter(Number.isFinite);
  if (matches.length >= 2) return (matches[0] + matches[1]) / 2;
  if (matches.length === 1) return matches[0];
  return parseRangeAverage(raw);
}

function inferTempScore(tempRequerida = "") {
  const raw = normalizeTempRiskText(tempRequerida);
  if (!raw) return 2;
  if (hasTempMarker(raw, "-60")) return 4;
  if (raw.includes("vivo") || hasTempMarker(raw, "7")) return 2;
  if (raw.includes("fresco") || hasTempMarker(raw, "4")) return 3;
  if (raw.includes("congelado") || hasTempMarker(raw, "-18")) return 1;
  return 2;
}

function inferLogisticRiskScore(riskText = "") {
  const normalized = normalizeGeoKey(riskText);
  if (normalized.includes("CRITICO")) return 4;
  if (normalized.includes("ALTO")) return 3;
  if (normalized.includes("MODERADO") || normalized.includes("MEDIO")) return 2;
  if (normalized.includes("BAJO")) return 1;
  return 2;
}

function inferCruceScore(empresa) {
  const cruce = normalizeGeoKey(`${empresa?.cruceFronterizo || ""} ${empresa?.rutaTerrestre || ""}`);
  if (cruce.includes("OTAY")) return 3;
  if (cruce.includes("NOGALES")) return 3;
  if (cruce.includes("NUEVO LAREDO") || cruce.includes("LAREDO")) return 2;
  if (cruce.includes("REYNOSA") || cruce.includes("MCALLEN") || cruce.includes("PHARR")) return 2;
  return 2;
}

function countCertificaciones(certText = "") {
  const raw = compactRouteText(certText);
  if (!raw || normalizeGeoKey(raw) === normalizeGeoKey(NO_INFO)) return 0;
  return raw
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean).length;
}

function buildClusterRecord(empresa) {
  const key = normalizeCompanyKey(empresa?.empresa || "");
  const override = CLUSTER_METRIC_OVERRIDES[key] || {};
  const baselineSales = Number(CLUSTER_SALES_BASELINE_MUSD[key]);
  const annualTrips = getEmpresaViajesAnuales(empresa);
  const duaMesFromAnnual = Number.isFinite(annualTrips) && annualTrips > 0 ? annualTrips / 12 : NaN;
  const duaMesFromCsv = parseFlexibleNumber(empresa?.duaMes || "");
  const ftlMes = Number.isFinite(override.ftlMes)
    ? override.ftlMes
    : Number.isFinite(duaMesFromAnnual)
      ? duaMesFromAnnual
      : Number.isFinite(duaMesFromCsv)
        ? duaMesFromCsv
        : parseRangeAverage(empresa?.volumenEstimado || "");
  const ventasFromCsv = parseSalesMUsd(empresa?.ventasAnuales || "");
  const ventasMUsd = Number.isFinite(override.ventasMUsd)
    ? override.ventasMUsd
    : Number.isFinite(ventasFromCsv)
      ? ventasFromCsv
      : Number.isFinite(baselineSales)
        ? baselineSales
        : 20;
  const tempScore = Number.isFinite(override.tempScore) ? override.tempScore : inferTempScore(empresa?.tempRequerida || "");
  const riesgoLogistico = Number.isFinite(override.riesgoLogistico)
    ? override.riesgoLogistico
    : inferLogisticRiskScore(empresa?.riesgoLogisticoCsv || "");
  const cruceScore = Number.isFinite(override.cruceScore) ? override.cruceScore : inferCruceScore(empresa);
  const numCertificaciones = Number.isFinite(override.numCertificaciones)
    ? override.numCertificaciones
    : countCertificaciones(empresa?.certificaciones || "");

  return {
    empresa: empresa?.empresa || "Sin empresa",
    ftlMes: Number.isFinite(ftlMes) ? ftlMes : 40,
    ventasMUsd: Number.isFinite(ventasMUsd) ? ventasMUsd : 20,
    tempScore,
    riesgoLogistico,
    cruceScore,
    numCertificaciones,
  };
}

function buildClusterFeatureMatrix(records = []) {
  const rows = records.map((r) => [
    Number(r.ftlMes) || 0,
    Number(r.ventasMUsd) || 0,
    Number(r.tempScore) || 0,
    Number(r.riesgoLogistico) || 0,
    Number(r.cruceScore) || 0,
    Number(r.numCertificaciones) || 0,
  ]);
  if (!rows.length) return [];

  const dims = rows[0].length;
  const means = Array.from({ length: dims }, (_, dim) => rows.reduce((acc, row) => acc + row[dim], 0) / rows.length);
  const stds = Array.from({ length: dims }, (_, dim) => {
    const variance = rows.reduce((acc, row) => acc + (row[dim] - means[dim]) ** 2, 0) / Math.max(1, rows.length - 1);
    return Math.sqrt(variance) || 1;
  });

  return rows.map((row) => row.map((value, dim) => (value - means[dim]) / stds[dim]));
}

function euclideanDistanceSq(a = [], b = []) {
  let sum = 0;
  for (let i = 0; i < Math.min(a.length, b.length); i += 1) {
    const d = a[i] - b[i];
    sum += d * d;
  }
  return sum;
}

function averageVectors(vectors = []) {
  if (!vectors.length) return [];
  const dims = vectors[0].length;
  const sum = Array(dims).fill(0);
  vectors.forEach((v) => {
    for (let i = 0; i < dims; i += 1) sum[i] += v[i];
  });
  return sum.map((v) => v / vectors.length);
}

function pickInitialKmeansCentroids(points = [], records = []) {
  if (!points.length) return [];
  const used = new Set();
  const indices = [];

  const addIndex = (idx) => {
    if (!Number.isInteger(idx) || idx < 0 || idx >= points.length || used.has(idx)) return false;
    used.add(idx);
    indices.push(idx);
    return true;
  };

  const topAnchorIdx = records
    .map((r, idx) => ({ idx, score: (Number(r.ftlMes) || 0) + (Number(r.ventasMUsd) || 0) }))
    .sort((a, b) => b.score - a.score)[0]?.idx;
  addIndex(topAnchorIdx);

  const topRiskIdx = records
    .map((r, idx) => ({ idx, score: (Number(r.ftlMes) || 0) * 0.6 + (Number(r.riesgoLogistico) || 0) * 0.4 }))
    .sort((a, b) => b.score - a.score)
    .find((entry) => !used.has(entry.idx))?.idx;
  addIndex(topRiskIdx);

  const minScaleIdx = records
    .map((r, idx) => ({ idx, score: (Number(r.ftlMes) || 0) + (Number(r.ventasMUsd) || 0) }))
    .sort((a, b) => a.score - b.score)
    .find((entry) => !used.has(entry.idx))?.idx;
  addIndex(minScaleIdx);

  for (let i = 0; i < points.length && indices.length < 3; i += 1) addIndex(i);
  return indices.map((idx) => [...points[idx]]);
}

function runKMeans(points = [], records = [], k = 3, maxIterations = 40) {
  if (!points.length) return { assignments: [], centroids: [], inertia: 0 };
  const safeK = Math.min(k, points.length);
  let centroids = pickInitialKmeansCentroids(points, records).slice(0, safeK);
  let assignments = Array(points.length).fill(0);

  for (let iter = 0; iter < maxIterations; iter += 1) {
    const nextAssignments = points.map((point) => {
      let bestIdx = 0;
      let bestDist = Number.POSITIVE_INFINITY;
      centroids.forEach((centroid, cIdx) => {
        const dist = euclideanDistanceSq(point, centroid);
        if (dist < bestDist) {
          bestDist = dist;
          bestIdx = cIdx;
        }
      });
      return bestIdx;
    });

    const changed = nextAssignments.some((value, idx) => value !== assignments[idx]);
    assignments = nextAssignments;

    const nextCentroids = centroids.map((centroid, cIdx) => {
      const members = points.filter((_, idx) => assignments[idx] === cIdx);
      return members.length ? averageVectors(members) : centroid;
    });
    centroids = nextCentroids;
    if (!changed) break;
  }

  const inertia = points.reduce((acc, point, idx) => {
    const centroid = centroids[assignments[idx]];
    return acc + euclideanDistanceSq(point, centroid);
  }, 0);

  return { assignments, centroids, inertia };
}

function covarianceMatrix(points = []) {
  if (!points.length) return [];
  const n = points.length;
  const d = points[0].length;
  const cov = Array.from({ length: d }, () => Array(d).fill(0));
  for (let i = 0; i < d; i += 1) {
    for (let j = i; j < d; j += 1) {
      let sum = 0;
      for (let r = 0; r < n; r += 1) sum += points[r][i] * points[r][j];
      const value = sum / Math.max(1, n - 1);
      cov[i][j] = value;
      cov[j][i] = value;
    }
  }
  return cov;
}

function matrixVectorMultiply(matrix = [], vector = []) {
  return matrix.map((row) => row.reduce((acc, value, idx) => acc + value * (vector[idx] || 0), 0));
}

function normalizeVector(vector = []) {
  const norm = Math.sqrt(vector.reduce((acc, value) => acc + value * value, 0)) || 1;
  return vector.map((value) => value / norm);
}

function powerIteration(matrix = [], iterations = 80, seedOffset = 0) {
  const dim = matrix.length;
  if (!dim) return { vector: [], eigenvalue: 0 };
  let vector = Array.from({ length: dim }, (_, idx) => (idx === seedOffset % dim ? 1 : 0.35));
  vector = normalizeVector(vector);

  for (let i = 0; i < iterations; i += 1) {
    const multiplied = matrixVectorMultiply(matrix, vector);
    vector = normalizeVector(multiplied);
  }

  const mv = matrixVectorMultiply(matrix, vector);
  const eigenvalue = vector.reduce((acc, value, idx) => acc + value * mv[idx], 0);
  return { vector, eigenvalue };
}

function deflateMatrix(matrix = [], eigenvector = [], eigenvalue = 0) {
  const dim = matrix.length;
  const next = matrix.map((row) => [...row]);
  for (let i = 0; i < dim; i += 1) {
    for (let j = 0; j < dim; j += 1) {
      next[i][j] -= eigenvalue * (eigenvector[i] || 0) * (eigenvector[j] || 0);
    }
  }
  return next;
}

function projectPca2D(points = []) {
  if (!points.length) return [];
  if (points[0].length < 2) return points.map((row) => ({ x: row[0] || 0, y: 0 }));

  const cov = covarianceMatrix(points);
  const first = powerIteration(cov, 80, 0);
  const covDeflated = deflateMatrix(cov, first.vector, first.eigenvalue);
  const second = powerIteration(covDeflated, 80, 1);

  return points.map((row) => ({
    x: row.reduce((acc, value, idx) => acc + value * (first.vector[idx] || 0), 0),
    y: row.reduce((acc, value, idx) => acc + value * (second.vector[idx] || 0), 0),
  }));
}

function spreadOverlappingPcaRows(rows = []) {
  if (!Array.isArray(rows) || !rows.length) return [];
  if (rows.length === 1) return rows.map((row) => ({ ...row }));

  const cloned = rows.map((row) => ({
    ...row,
    pcaX: Number(row.pcaX) || 0,
    pcaY: Number(row.pcaY) || 0,
  }));

  const getSpans = () => {
    const xs = cloned.map((row) => row.pcaX);
    const ys = cloned.map((row) => row.pcaY);
    return {
      xSpan: Math.max(0.001, Math.max(...xs) - Math.min(...xs)),
      ySpan: Math.max(0.001, Math.max(...ys) - Math.min(...ys)),
    };
  };

  const { xSpan, ySpan } = getSpans();
  const minDx = Math.max(0.05, xSpan * 0.08);
  const minDy = Math.max(0.05, ySpan * 0.08);
  const baseJitterX = Math.max(0.01, xSpan * 0.02);
  const baseJitterY = Math.max(0.01, ySpan * 0.02);

  // Jitter determinista inicial para evitar solapes exactos.
  cloned.forEach((row, idx) => {
    const angle = (2 * Math.PI * idx) / cloned.length;
    row.pcaX += Math.cos(angle) * baseJitterX * 0.35;
    row.pcaY += Math.sin(angle) * baseJitterY * 0.35;
  });

  // Separación iterativa para puntos demasiado cercanos en pantalla.
  for (let iter = 0; iter < 10; iter += 1) {
    let moved = false;
    for (let i = 0; i < cloned.length; i += 1) {
      for (let j = i + 1; j < cloned.length; j += 1) {
        const a = cloned[i];
        const b = cloned[j];
        const dx = b.pcaX - a.pcaX;
        const dy = b.pcaY - a.pcaY;
        const closeX = Math.abs(dx) < minDx;
        const closeY = Math.abs(dy) < minDy;
        if (!closeX || !closeY) continue;

        const pushX = (minDx - Math.abs(dx)) * 0.5;
        const pushY = (minDy - Math.abs(dy)) * 0.5;
        const dirX = dx >= 0 ? 1 : -1;
        const dirY = dy >= 0 ? 1 : -1;

        a.pcaX -= dirX * pushX;
        b.pcaX += dirX * pushX;
        a.pcaY -= dirY * pushY;
        b.pcaY += dirY * pushY;
        moved = true;
      }
    }
    if (!moved) break;
  }

  return cloned;
}

function profileByCluster(assignments = [], records = []) {
  const stats = new Map();
  assignments.forEach((clusterId, idx) => {
    if (!stats.has(clusterId)) stats.set(clusterId, { members: [], avgFtl: 0, avgVentas: 0, avgRiesgo: 0 });
    stats.get(clusterId).members.push(records[idx]);
  });

  stats.forEach((clusterStat) => {
    const members = clusterStat.members;
    clusterStat.avgFtl = members.reduce((acc, r) => acc + (Number(r.ftlMes) || 0), 0) / Math.max(1, members.length);
    clusterStat.avgVentas = members.reduce((acc, r) => acc + (Number(r.ventasMUsd) || 0), 0) / Math.max(1, members.length);
    clusterStat.avgRiesgo =
      members.reduce((acc, r) => acc + (Number(r.riesgoLogistico) || 0), 0) / Math.max(1, members.length);
  });

  const clusters = Array.from(stats.keys());
  const labelMap = {};
  if (!clusters.length) return labelMap;

  const anclaCluster = clusters
    .map((clusterId) => ({ clusterId, score: stats.get(clusterId).avgFtl + stats.get(clusterId).avgVentas }))
    .sort((a, b) => b.score - a.score)[0]?.clusterId;
  labelMap[anclaCluster] = "Ancla";

  const remaining = clusters.filter((clusterId) => clusterId !== anclaCluster);
  if (remaining.length) {
    const estrategicoCluster = remaining
      .map((clusterId) => ({
        clusterId,
        score: stats.get(clusterId).avgFtl * 0.6 + stats.get(clusterId).avgRiesgo * 0.4,
      }))
      .sort((a, b) => b.score - a.score)[0]?.clusterId;
    if (estrategicoCluster !== undefined) labelMap[estrategicoCluster] = "Estratégico";
    remaining.forEach((clusterId) => {
      if (!labelMap[clusterId]) labelMap[clusterId] = "Estratégico";
    });
  }

  return labelMap;
}

function getClusteredProspectData() {
  const selectedCompanies = selectPriorityClusterCompanies(empresasData);
  const records = selectedCompanies.map((empresa) => buildClusterRecord(empresa));
  const recordsByKey = new Map(records.map((record) => [normalizeCompanyKey(record.empresa), record]));

  const rowsWithFeatures = CLUSTER_BASE_ROWS.map((item) => {
    const record =
      recordsByKey.get(normalizeCompanyKey(item.empresa)) ||
      records.find((r) => normalizeCompanyKey(r.empresa).includes(normalizeCompanyKey(item.empresa)));
    return {
      ...item,
      ftlMesFeature: Number.isFinite(record?.ftlMes) ? record.ftlMes : item.ftlValue,
      ventasMUsd: Number.isFinite(record?.ventasMUsd) ? record.ventasMUsd : 0,
      tempScore: Number.isFinite(record?.tempScore) ? record.tempScore : 2,
      riesgoLogisticoScore: Number.isFinite(record?.riesgoLogistico)
        ? record.riesgoLogistico
        : inferLogisticRiskScore(item.riesgo || ""),
      cruceScore: Number.isFinite(record?.cruceScore) ? record.cruceScore : 2,
      numCertificaciones: Number.isFinite(record?.numCertificaciones) ? record.numCertificaciones : 0,
    };
  });

  const clusteringRecords = rowsWithFeatures.map((item) => ({
    ftlMes: Number(item.ftlMesFeature) || 0,
    ventasMUsd: Number(item.ventasMUsd) || 0,
    tempScore: Number(item.tempScore) || 0,
    riesgoLogistico: Number(item.riesgoLogisticoScore) || 0,
    cruceScore: Number(item.cruceScore) || 0,
    numCertificaciones: Number(item.numCertificaciones) || 0,
  }));
  const clusteringMatrix = buildClusterFeatureMatrix(clusteringRecords);
  const clusteringResult = runKMeans(clusteringMatrix, clusteringRecords, 2);
  const profileMap = profileByCluster(clusteringResult.assignments, clusteringRecords);
  const rows = rowsWithFeatures.map((item, idx) => {
    const clusterId = Number.isInteger(clusteringResult.assignments[idx]) ? clusteringResult.assignments[idx] : 0;
    return {
      ...item,
      clusterId,
      profile: profileMap[clusterId] || "Estratégico",
    };
  });

  return {
    rows,
    clusteringRecords,
    clusteringMatrix,
    recordsCount: records.length,
  };
}

function getClusteredProspectRowByName(name = "") {
  const key = normalizeCompanyKey(name);
  if (!key) return null;
  const { rows } = getClusteredProspectData();
  return (
    rows.find((row) => normalizeCompanyKey(row.empresa) === key) ||
    rows.find((row) => normalizeCompanyKey(row.empresa).includes(key) || key.includes(normalizeCompanyKey(row.empresa))) ||
    null
  );
}

function clusterRiskLabel(score) {
  if (score >= 4) return "Crítico";
  if (score >= 3) return "Alto";
  if (score >= 2) return "Moderado";
  return "Bajo";
}

function clusterQuarterByProfile(profile) {
  if (profile === "Ancla") return "Q1 2026";
  if (profile === "Estratégico") return "Q2 2026";
  return "Q3 2026";
}

function forcedClusterByCompanyName(name = "") {
  const key = normalizeCompanyKey(name);
  if (key.includes("GRUPO PINSA")) return 0;
  if (key.includes("GRUPO ACUICOLA MEXICANO") || key === "GAM" || key.includes(" GAM")) return 0;
  if (key.includes("BAJA AQUA FARMS")) return 2;
  if (key.includes("PACIFICO AQUACULTURE")) return 1;
  if (key.includes("BAJA SHELLFISH FARMS")) return 1;
  return null;
}

function clusterProfileById(clusterId) {
  if (clusterId === 0) return "Ancla";
  if (clusterId === 2) return "Estratégico";
  return "Estratégico";
}

function forcedClusterDisplayOrder(name = "") {
  const key = normalizeCompanyKey(name);
  if (key.includes("GRUPO PINSA")) return 0;
  if (key.includes("GRUPO ACUICOLA MEXICANO") || key === "GAM" || key.includes(" GAM")) return 1;
  if (key.includes("BAJA AQUA FARMS")) return 2;
  if (key.includes("PACIFICO AQUACULTURE")) return 3;
  if (key.includes("BAJA SHELLFISH FARMS")) return 4;
  return 99;
}

function renderClustering() {
  const plotEl = document.getElementById("clusterScatterPlot");
  const pcaEl = document.getElementById("clusterPcaPlot");
  const elbowEl = document.getElementById("clusterElbowPlot");
  const tbody = document.getElementById("clusterSummaryBody");
  const noticeEl = document.getElementById("clusterNotice");
  if (!plotEl || !elbowEl || !tbody) return;
  setClusterPcaVisibility();
  const showPca = !!pcaVisible;
  const clusterData = getClusteredProspectData();
  const fixedRows = clusterData.rows;
  const clusteringRecords = clusterData.clusteringRecords;
  const clusteringMatrix = clusterData.clusteringMatrix;

  if (clusterData.recordsCount < 3) {
    if (noticeEl) noticeEl.textContent = "Sin datos suficientes para clustering (mínimo 3 empresas).";
    tbody.innerHTML = '<tr><td colspan="6">Sin datos suficientes.</td></tr>';
    if (typeof Plotly !== "undefined") {
      Plotly.purge(plotEl);
      if (pcaEl) Plotly.purge(pcaEl);
      Plotly.purge(elbowEl);
    }
    return;
  }

  const profileOrder = ["Ancla", "Estratégico"];
  const profileColors = {
    Ancla: "#1a6b3a",
    Estratégico: "#8fcda2",
  };

  if (typeof Plotly !== "undefined") {
    const traces = profileOrder.map((profile) => {
      const items = fixedRows.filter((item) => item.profile === profile);
      const clusterRef = items.length ? items[0].clusterId : "";
      return {
        type: "scatter",
        mode: "markers",
        name: clusterRef === "" ? profile : `${profile} (Cluster ${clusterRef})`,
        x: items.map((item) => item.ftlValue),
        y: items.map((item) => item.ventasMUsd),
        customdata: items.map((item) => [item.empresa, item.profile, item.clusterId, item.ftlLabel, item.aduana, item.riesgo]),
        marker: {
          size: 13,
          color: profileColors[profile],
          line: { color: "#ffffff", width: 1.2 },
        },
        hovertemplate:
          "Empresa: %{customdata[0]}<br>Perfil: %{customdata[1]} (Cluster %{customdata[2]})<br>DUA/mes: %{customdata[3]}<br>Aduana: %{customdata[4]}<br>Riesgo: %{customdata[5]}<extra></extra>",
      };
    });

    Plotly.react(
      plotEl,
      traces,
      {
        title: { text: "", font: { family: "Montserrat, sans-serif", size: 16, color: "#1f3443" } },
        margin: { l: 56, r: 20, t: 24, b: 56 },
        paper_bgcolor: "#ffffff",
        plot_bgcolor: "#ffffff",
        hovermode: "closest",
        font: { family: "Montserrat, sans-serif", color: "#1f3443", size: 12 },
        legend: { orientation: "h", x: 0, y: 1, xanchor: "left", yanchor: "top", font: { size: 11, color: "#4b6475" } },
        xaxis: {
          title: "DUA/mes",
          tickfont: { color: "#355264" },
          showgrid: true,
          gridcolor: "rgba(10, 45, 74, 0.08)",
          zeroline: false,
        },
        yaxis: {
          title: "Ventas estimadas (MUSD)",
          tickfont: { color: "#355264" },
          gridcolor: "rgba(10, 45, 74, 0.10)",
          zeroline: false,
        },
      },
      {
        responsive: true,
        displaylogo: false,
        scrollZoom: false,
        modeBarButtonsToRemove: ["lasso2d", "select2d", "autoScale2d", "toggleSpikelines"],
      },
    );

    if (pcaEl && showPca) {
      const pcaFeatures = fixedRows.map((item) => ({
        ftlMes: Number(item.ftlMesFeature) || 0,
        ventasMUsd: Number(item.ventasMUsd) || 0,
        tempScore: Number(item.tempScore) || 0,
        riesgoLogistico: Number(item.riesgoLogisticoScore) || 0,
        cruceScore: Number(item.cruceScore) || 0,
        numCertificaciones: Number(item.numCertificaciones) || 0,
      }));
      const pcaMatrix = buildClusterFeatureMatrix(pcaFeatures);
      const pcaProjection = projectPca2D(pcaMatrix);
      const pcaRowsBase = fixedRows.map((item, idx) => ({
        ...item,
        pcaX: Number.isFinite(pcaProjection[idx]?.x) ? pcaProjection[idx].x : 0,
        pcaY: Number.isFinite(pcaProjection[idx]?.y) ? pcaProjection[idx].y : 0,
      }));
      const pcaRows = spreadOverlappingPcaRows(pcaRowsBase);
      const pcaXVals = pcaRows.map((item) => Number(item.pcaX) || 0);
      const pcaYVals = pcaRows.map((item) => Number(item.pcaY) || 0);
      const pcaXMin = Math.min(...pcaXVals);
      const pcaXMax = Math.max(...pcaXVals);
      const pcaYMin = Math.min(...pcaYVals);
      const pcaYMax = Math.max(...pcaYVals);
      const pcaXPad = Math.max(0.2, (pcaXMax - pcaXMin || 1) * 0.15);
      const pcaYPad = Math.max(0.2, (pcaYMax - pcaYMin || 1) * 0.15);

      const pcaTraces = profileOrder.map((profile) => {
        const items = pcaRows.filter((item) => item.profile === profile);
        return {
          type: "scatter",
          mode: "markers",
          name: profile === "Ancla" ? "Ancla (PCA)" : "Estratégico (PCA)",
          x: items.map((item) => item.pcaX),
          y: items.map((item) => item.pcaY),
          customdata: items.map((item) => [
            item.empresa,
            item.ftlLabel,
            item.ventasMUsd,
            item.tempScore,
            item.riesgoLogisticoScore,
            item.cruceScore,
            item.numCertificaciones,
          ]),
          marker: {
            size: 10,
            color: profileColors[profile],
            line: { color: "#ffffff", width: 1.2 },
          },
          hovertemplate:
            "Empresa: %{customdata[0]}<br>DUA/mes: %{customdata[1]}<br>Ventas est.: %{customdata[2]:.1f} MUSD<br>Temp score: %{customdata[3]}<br>Riesgo logístico: %{customdata[4]}<br>Cruce score: %{customdata[5]}<br>Certificaciones: %{customdata[6]}<extra></extra>",
        };
      });

      Plotly.react(
        pcaEl,
        pcaTraces,
        {
          title: {
            text: "Proyección PCA 2D (6 features del clustering)",
            x: 0.5,
            y: 0.97,
            font: { family: "Montserrat, sans-serif", size: 15, color: "#1f3443" },
          },
          margin: { t: 60, b: 50, l: 60, r: 20 },
          paper_bgcolor: "#ffffff",
          plot_bgcolor: "#ffffff",
          hovermode: "closest",
          font: { family: "Montserrat, sans-serif", color: "#1f3443", size: 12 },
          showlegend: false,
          xaxis: {
            title: "Componente principal 1",
            range: [pcaXMin - pcaXPad, pcaXMax + pcaXPad],
            tickfont: { color: "#355264" },
            showgrid: true,
            gridcolor: "rgba(10, 45, 74, 0.08)",
            zeroline: false,
          },
          yaxis: {
            title: "Componente principal 2",
            range: [pcaYMin - pcaYPad, pcaYMax + pcaYPad],
            tickfont: { color: "#355264" },
            gridcolor: "rgba(10, 45, 74, 0.10)",
            zeroline: false,
          },
        },
        {
          responsive: true,
          displaylogo: false,
          scrollZoom: false,
          modeBarButtonsToRemove: ["lasso2d", "select2d", "autoScale2d", "toggleSpikelines"],
        },
      );
    }
    if (pcaEl && !showPca) {
      Plotly.purge(pcaEl);
    }

    const elbowKs = [1, 2, 3, 4].filter((k) => k <= Math.max(1, clusteringRecords.length));
    const elbowVals = elbowKs.map((k) => runKMeans(clusteringMatrix, clusteringRecords, k).inertia);
    const hasSecondK = elbowVals.length > 1;
    const dropPercent = hasSecondK && elbowVals[0] > 0 ? ((elbowVals[0] - elbowVals[1]) / elbowVals[0]) * 100 : 0;
    const annotationX = hasSecondK ? elbowKs[1] : elbowKs[0];
    const annotationY = hasSecondK ? elbowVals[1] : elbowVals[0];
    Plotly.react(
      elbowEl,
      [
        {
          type: "scatter",
          mode: "lines+markers+text",
          x: elbowKs,
          y: elbowVals,
          text: elbowVals.map((v) => v.toFixed(2)),
          textposition: "top center",
          line: { color: "#046f31", width: 3 },
          marker: { color: "#046f31", size: 8 },
          hovertemplate: "k=%{x}<br>Inercia=%{y:.2f}<extra></extra>",
          name: "Inercia",
        },
      ],
      {
        title: {
          text: "Elbow Graph — Selección de k óptimo",
          font: { family: "Montserrat, sans-serif", size: 16, color: "#1f3443" },
        },
        margin: { l: 66, r: 22, t: 56, b: 56 },
        paper_bgcolor: "#ffffff",
        plot_bgcolor: "#ffffff",
        hovermode: "closest",
        font: { family: "Montserrat, sans-serif", color: "#1f3443", size: 12 },
        showlegend: false,
        xaxis: {
          title: "Número de clusters (k)",
          tickmode: "array",
          tickvals: elbowKs,
          ticktext: elbowKs.map(String),
          showgrid: false,
          zeroline: false,
        },
        yaxis: {
          title: "Inercia",
          gridcolor: "rgba(10, 45, 74, 0.10)",
          zeroline: false,
        },
        annotations: [
          {
            x: annotationX,
            y: annotationY,
            xanchor: "left",
            yanchor: "bottom",
            text: `k óptimo seleccionado<br>Mayor caída: -${dropPercent.toFixed(1)}%`,
            showarrow: true,
            arrowhead: 2,
            arrowsize: 1,
            arrowwidth: 1.2,
            arrowcolor: "#1f3443",
            ax: 90,
            ay: -34,
            bgcolor: "rgba(255,255,255,0.95)",
            bordercolor: "#d7e3ec",
            borderwidth: 1,
            font: { family: "Montserrat, sans-serif", size: 11, color: "#1f3443" },
          },
        ],
      },
      {
        responsive: true,
        displaylogo: false,
        scrollZoom: false,
        modeBarButtonsToRemove: ["lasso2d", "select2d", "autoScale2d", "toggleSpikelines"],
      },
    );
  } else {
    if (noticeEl) noticeEl.textContent = "No se pudo cargar Plotly para clustering.";
  }

  if (noticeEl) noticeEl.textContent = "";
  tbody.innerHTML = fixedRows
    .map(
      (item) => `
      <tr>
        <td>${escapeHtml(item.empresa)}</td>
        <td>${escapeHtml(item.profile)}</td>
        <td>${escapeHtml(item.ftlLabel)}</td>
        <td>${escapeHtml(item.aduana)}</td>
        <td>${escapeHtml(item.riesgo)}</td>
        <td>${item.quarter}</td>
      </tr>
    `,
    )
    .join("");

  if (typeof Plotly !== "undefined") {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        [plotEl, pcaEl, elbowEl].forEach((el) => {
          if (el && el.data) Plotly.Plots.resize(el);
        });
      });
    });
  }
}

function infraRiskLevelToScore(level) {
  const normalized = normalizeInfraRiskLevel(level);
  if (normalized === "CRITICO") return 4;
  if (normalized === "ALTO") return 3;
  return 2;
}

function scoreToRiskLevel(score) {
  if (score >= 3.5) return "critico";
  if (score >= 2.6) return "alto";
  if (score >= 1.8) return "medio";
  return "bajo";
}

function riskLevelTitle(level) {
  const normalized = normalizeRiskLevel(level);
  if (normalized === "critico") return "Crítico";
  if (normalized === "alto") return "Alto";
  if (normalized === "medio") return "Medio";
  return "Bajo";
}

function buildPropuestaProspect(empresa) {
  const rutas = resolveEmpresaRutas(empresa);
  const rutaTerrestre = rutas.terrestre;
  const rutaTerrestreRaw = getEmpresaRutaTerrestreRaw(empresa, rutaTerrestre.nombre);
  const rutaTerrestreCsv =
    detectTerrestreByText(normalizeGeoKey(`${empresa.cruceFronterizo || ""} ${rutaTerrestreRaw || ""}`)) ||
    rutaTerrestre;
  const cruceInfo =
    findCruceByEmpresa(
      empresa,
      `${rutaTerrestreRaw || ""} | ${rutaTerrestreCsv?.nombre || ""} | ${rutaTerrestre?.nombre || ""}`,
    ) ||
    findCruceByRutaReferencia(
      `${rutaTerrestreCsv?.nombre || ""} | ${empresa?.cruceFronterizo || ""} | ${rutaTerrestreRaw || ""}`,
    );

  const logRisk = parseEmpresaLogRiskCsv(empresa.riesgoLogisticoCsv || "");
  const logScore = riskLevelScore(logRisk?.level || "medio");
  const aduanaScore = infraRiskLevelToScore(cruceInfo?.riesgoLevel || "MODERADO");
  const tempReq = empresa.tempRequerida || inferTempRequerida(empresa.productos || empresa.especialidad || "", empresa.actividad || "");
  const thresholdC = getCriticalThreshold(tempReq);
  let coldSensitivityScore = 1;
  if (thresholdC <= 25) coldSensitivityScore = 4;
  else if (thresholdC <= 28) coldSensitivityScore = 3;
  else if (thresholdC <= 32) coldSensitivityScore = 2;

  const score = logScore * 0.5 + aduanaScore * 0.3 + coldSensitivityScore * 0.2;
  const level = scoreToRiskLevel(score);
  const aduana = getAduanaDisplayName(cruceInfo);

  return {
    empresa: empresa.empresa || "Sin empresa",
    aduana,
    score,
    level,
  };
}

function renderPropuestaTab() {
  const fobEl = document.getElementById("propuestaHeroFobValue");
  if (fobEl) {
    fobEl.textContent = "$720M USD";
  }
  ensurePropuestaValorSection();
  syncPropuestaProspectsLocations();
  syncPropuestaPlanClusters();
  initPropuestaCoverageMap();
  bindPropuestaProspectButtons();
  bindPropuestaViabilidadButton();
}

function ensurePropuestaValorStyles() {
  if (document.getElementById("propuestaValorStyles")) return;
  const style = document.createElement("style");
  style.id = "propuestaValorStyles";
  style.textContent = `
    .propuesta-valor-coverage-wrap{
      display:grid;
      grid-template-columns:minmax(0,1fr) minmax(0,1fr);
      gap:0.72rem;
      align-items:stretch;
    }
    .propuesta-valor-coverage-wrap .propuesta-section{
      margin:0;
    }
    .propuesta-valor-grid{
      display:grid;
      grid-template-columns:repeat(3,minmax(0,1fr));
      gap:0.52rem;
      align-items:stretch;
    }
    @media (max-width:1100px){
      .propuesta-valor-coverage-wrap{
        grid-template-columns:1fr;
      }
    }
    @media (max-width:900px){
      .propuesta-valor-grid{grid-template-columns:1fr}
    }
  `;
  document.head.appendChild(style);
}

function buildPropuestaValorHtml() {
  return `
    <h3 class="propuesta-section-title">Propuesta de valor</h3>
    <div class="propuesta-valor-grid">
      <div class="propuesta-implementation-card propuesta-plan-phase">
        <h4>📋 Cumplimiento FSMA 204</h4>
        <p>Cada viaje genera un registro térmico certificado por ENAC — la documentación que el importador en EE.UU. exige desde enero 2026.</p>
      </div>
      <div class="propuesta-implementation-card propuesta-plan-phase">
        <h4>⚠️ Mitigación de riesgo real</h4>
        <p>Un rechazo en frontera supera el costo anual del servicio. CLCircular convierte ese riesgo en un certificado automático por viaje.</p>
      </div>
      <div class="propuesta-implementation-card propuesta-plan-phase">
        <h4>🔄 Modelo sin fricción</h4>
        <p>Sin inversión inicial. Sin configuración. Pago por viaje monitoreado — el exportador paga por resultado, no por hardware.</p>
      </div>
    </div>
  `;
}

function ensurePropuestaValorSection() {
  const tab = document.getElementById("tab-propuesta");
  if (!tab) return;
  ensurePropuestaValorStyles();
  const planSection = tab.querySelector(".propuesta-plan");
  const coverageSection = tab.querySelector(".propuesta-coverage");
  let valorSection = document.getElementById("propuestaValorSection");
  let wrap = document.getElementById("propuestaValorCoverageWrap");

  if (!valorSection) {
    valorSection = document.createElement("article");
    valorSection.id = "propuestaValorSection";
    valorSection.className = "panel propuesta-section propuesta-valor";
    valorSection.innerHTML = buildPropuestaValorHtml();
  } else {
    valorSection.innerHTML = buildPropuestaValorHtml();
  }

  if (coverageSection) {
    if (!wrap) {
      wrap = document.createElement("div");
      wrap.id = "propuestaValorCoverageWrap";
      wrap.className = "propuesta-valor-coverage-wrap";
    }
    if (planSection) {
      tab.insertBefore(wrap, planSection);
    } else if (!wrap.parentElement) {
      tab.appendChild(wrap);
    }
    wrap.appendChild(coverageSection);
    wrap.appendChild(valorSection);
    return;
  }

  if (planSection) {
    tab.insertBefore(valorSection, planSection);
  } else if (!valorSection.parentElement) {
    tab.appendChild(valorSection);
  }
}

function getViabilidadScenarioData() {
  const current = state.viabilidadScenario;
  if (current === "Conservador") return VIAB_PL_CONS;
  if (current === "Optimista") return VIAB_PL_OPT;
  return VIAB_PL_BASE;
}

function viabFormatMoney(value) {
  return `$${Math.abs(Number(value) || 0).toLocaleString("es-MX")}`;
}

function viabFormatK(value) {
  const num = Math.round(Math.abs(Number(value) || 0) / 1000);
  return `$${num.toLocaleString("es-MX")}K`;
}

function viabFormatSignedK(value) {
  const num = Number(value) || 0;
  const absK = Math.round(Math.abs(num) / 1000).toLocaleString("es-MX");
  return `${num < 0 ? "-" : ""}$${absK}K`;
}

function viabFormatInt(value) {
  return Math.round(Math.abs(Number(value) || 0)).toLocaleString("es-MX");
}

function viabRiskClass(level = "") {
  const normalized = normalizeHeader(level).toUpperCase();
  if (normalized.includes("CRIT")) return "critico";
  if (normalized.includes("ALTO")) return "alto";
  if (normalized.includes("MOD")) return "moderado";
  return "moderado";
}

function viabScenarioButtonsHtml(current) {
  return ["Base", "Conservador", "Optimista"]
    .map(
      (label) =>
        `<button class="viab-scenario-btn${label === current ? " is-active" : ""}" type="button" data-viab-scenario="${label}">${label}</button>`,
    )
    .join("");
}

function viabRenderPnlTableRows(data) {
  const rows = [
    { label: "Revenue total", key: "rev", fmt: (v) => viabFormatK(v) },
    { label: "(-) COGS", key: "cogs", fmt: (v) => `(${viabFormatK(v)})` },
    { label: "Utilidad Bruta", custom: (d) => viabFormatK((d.rev || 0) - (d.cogs || 0)), highlight: true },
    { label: "(-) OPEX", key: "opex", fmt: (v) => `(${viabFormatK(v)})` },
    { label: "EBITDA", key: "ebitda", fmt: (v) => viabFormatK(v), highlight: true },
    { label: "Margen EBITDA", key: "ebm", fmt: (v) => `${Number(v || 0).toFixed(1)}%` },
    {
      label: "Utilidad Neta",
      key: "ni",
      fmt: (v) => {
        const num = Number(v) || 0;
        return num >= 0 ? viabFormatK(num) : `(${viabFormatK(Math.abs(num))})`;
      },
      highlight: true,
      negative: true,
    },
    { label: "Margen Neto", key: "nm", fmt: (v) => `${Number(v || 0).toFixed(1)}%` },
    { label: "Viajes", key: "viajes", fmt: (v) => viabFormatInt(v) },
  ];

  return rows
    .map((row) => {
      const cells = data
        .map((item) => {
          const raw = row.custom ? row.custom(item) : row.fmt(item[row.key]);
          const negClass = row.negative && Number(item[row.key]) < 0 ? " viab-neg" : "";
          return `<td class="${negClass.trim()}">${raw}</td>`;
        })
        .join("");
      return `<tr class="${row.highlight ? "viab-highlight-row" : ""}"><td>${row.label}</td>${cells}</tr>`;
    })
    .join("");
}

function viabRenderClientsYear1() {
  return VIAB_CLIENTS_Y1.map((c) => {
    const riskColor = VIAB_RISK_COLORS[c.risk] || VIAB_COLORS.yellow;
    return `
      <div class="viab-client-card" style="border-left-color:${riskColor}">
        <div class="viab-client-head">
          <div>
            <strong>${escapeHtml(c.name)}</strong>
            <span class="viab-tier-badge">${escapeHtml(c.tier)}</span>
            <span class="viab-risk-badge viab-risk-${viabRiskClass(c.risk)}">${escapeHtml(c.risk === "CRITICO" ? "CRÍTICO" : c.risk)}</span>
          </div>
          <span class="viab-client-quarter">${escapeHtml(c.q)}</span>
        </div>
        <div class="viab-client-meta">${viabFormatInt(c.mes)} DUA/mes · ${escapeHtml(c.sede)} · ${escapeHtml(c.cruce)}</div>
        <div class="viab-client-foot">
          <span>$${viabFormatInt(c.precio)}/viaje · ${viabFormatInt(c.devices)} sensores</span>
          <strong>${viabFormatMoney(c.rev)}/año</strong>
        </div>
      </div>
    `;
  }).join("");
}

function viabRenderClientsYear2() {
  return VIAB_CLIENTS_Y2.map(
    (c) => `
      <div class="viab-client-card is-muted">
        <div class="viab-client-head">
          <div>
            <strong>${escapeHtml(c.name)}</strong>
            <span class="viab-tier-badge">${escapeHtml(c.tier)}</span>
          </div>
          <strong>${viabFormatMoney(c.rev)}/año</strong>
        </div>
        <div class="viab-client-meta">${viabFormatInt(c.mes)} DUA/mes · $${viabFormatInt(c.precio)}/viaje</div>
      </div>
    `,
  ).join("");
}

function viabRenderMembershipBars() {
  const maxRev = Math.max(...VIAB_CLIENTS_Y1.map((c) => c.rev));
  return VIAB_CLIENTS_Y1.map((c) => {
    const tierColor = VIAB_COLORS.dark;
    const width = maxRev > 0 ? Math.round((c.rev / maxRev) * 100) : 0;
    const margin = ((c.precio - 5.17) / c.precio) * 100;
    return `
      <div class="viab-membership-bar-row">
        <div class="viab-membership-bar-head">
          <span>
            <strong>${escapeHtml(c.name)}</strong>
            <span class="viab-chip" style="color:${tierColor};border-color:${tierColor}44;background:${tierColor}18">
              ${escapeHtml(c.tier)} · $${viabFormatInt(c.precio)}/viaje
            </span>
          </span>
          <strong>${viabFormatMoney(c.rev)}</strong>
        </div>
        <div class="viab-membership-track">
          <div class="viab-membership-fill" style="width:${width}%;background:${tierColor}"></div>
        </div>
        <div class="viab-membership-caption">${viabFormatInt(c.viajes)} viajes/año · ${viabFormatInt(c.mes)} DUA/mes · margen bruto ${margin.toFixed(1)}%</div>
      </div>
    `;
  }).join("");
}

function viabRenderInvestmentBars() {
  return VIAB_INVESTMENT.map(
    (item) => `
      <div class="viab-progress-row">
        <div class="viab-progress-head">
          <span>${escapeHtml(item.label)}</span>
          <span class="${item.warn ? "is-warn" : ""}">${viabFormatMoney(item.monto)} <small>(${item.pct}%)</small></span>
        </div>
        <div class="viab-progress-track">
          <div class="viab-progress-fill ${item.warn ? "is-warn" : ""}" style="width:${item.pct}%"></div>
        </div>
      </div>
    `,
  ).join("");
}

function viabRenderDrawdownBars() {
  const drawdowns = [
    { mes: "Mes 1", desc: "Legal + Capital de trabajo + Mktg", monto: 31000, pct: 40 },
    { mes: "Mes 2", desc: "Personal + Hub Mazatlan", monto: 37000, pct: 48 },
    { mes: "Mes 3", desc: "Flota sensores (256 x $35)", monto: 8960, pct: 12 },
  ];

  return drawdowns
    .map(
      (t) => `
      <div class="viab-progress-row">
        <div class="viab-progress-head">
          <span><strong>${t.mes}</strong> · ${t.desc}</span>
          <strong>${viabFormatMoney(t.monto)}</strong>
        </div>
        <div class="viab-progress-track">
          <div class="viab-progress-fill" style="width:${t.pct}%"></div>
        </div>
      </div>
    `,
    )
    .join("");
}

function viabRenderCostoNetoRows() {
  const rows = VIAB_CLIENTS_Y1.map((c) => {
    const recup = Math.round(c.rev * 0.51);
    const neto = c.rev - recup;
    const npv = neto / c.viajes;
    const ahorro = Math.round((27.5 - npv) * c.viajes);
    return `
      <tr>
        <td>${escapeHtml(c.name.split(" ")[0])}</td>
        <td>${viabFormatMoney(c.rev)}</td>
        <td class="viab-positive">${viabFormatMoney(recup)}</td>
        <td class="viab-strong">$${npv.toFixed(2)}/viaje</td>
        <td class="viab-positive">+${viabFormatMoney(ahorro)}</td>
      </tr>
    `;
  }).join("");

  return `${rows}<tr class="viab-highlight-row"><td>TOTAL</td><td>$98,515</td><td>$50,243</td><td>$48,272</td><td>Argumento de ventas</td></tr>`;
}

function viabRenderScenariosCards() {
  const cards = [
    {
      name: "Base",
      color: VIAB_COLORS.dark,
      rev: "$285K",
      ebm: "61.3%",
      ni: "$73K",
      cum: "$958K",
      pb: "9 meses",
      desc: "3% crec. · retención 100% · pérdida 5%",
    },
    {
      name: "Conservador",
      color: VIAB_COLORS.orange,
      rev: "$198K",
      ebm: "47.9%",
      ni: "$38K",
      cum: "$430K",
      pb: "18 meses",
      desc: "1.5% crec. · churn 20% · costos ×1.15",
    },
    {
      name: "Optimista",
      color: VIAB_COLORS.mid,
      rev: "$398K",
      ebm: "64.8%",
      ni: "$95K",
      cum: "$1.5M",
      pb: "7 meses",
      desc: "5% crec. · retención 100% · costos ×0.90",
    },
  ];

  return cards
    .map(
      (s) => `
      <div class="viab-scenario-card" style="border-top-color:${s.color}">
        <div class="viab-scenario-name" style="color:${s.color}">${s.name}</div>
        <div class="viab-scenario-kv"><span>Revenue Año 10</span><strong style="color:${s.color}">${s.rev}</strong></div>
        <div class="viab-scenario-kv"><span>Margen EBITDA</span><strong style="color:${s.color}">${s.ebm}</strong></div>
        <div class="viab-scenario-kv"><span>Utilidad Neta</span><strong style="color:${s.color}">${s.ni}</strong></div>
        <div class="viab-scenario-kv"><span>FCF Acumulado</span><strong style="color:${s.color}">${s.cum}</strong></div>
        <div class="viab-scenario-kv"><span>Payback</span><strong style="color:${s.color}">${s.pb}</strong></div>
        <p class="viab-scenario-desc">${s.desc}</p>
      </div>
    `,
    )
    .join("");
}

function viabRenderIndicadoresSection() {
  const irrValue = Number.isFinite(VIAB_IRR_BASE) ? `${(VIAB_IRR_BASE * 100).toFixed(1)}%` : "N/D";
  const npvNafinPositive = VIAB_NPV_NAFIN >= 0;
  const rentabilidad = [
    {
      l: "TIR (IRR)",
      v: irrValue,
      s: "Sobre FCF 10 años",
      badge: "EXCELENTE",
      bc: VIAB_COLORS.dark,
      color: VIAB_COLORS.dark,
      note: "Supera ampliamente el costo de capital (16%).",
    },
    {
      l: "VPN @ WACC 16%",
      v: viabFormatSignedK(VIAB_NPV_NAFIN),
      s: "Tasa NAFIN",
      badge: npvNafinPositive ? "POSITIVO" : "NEGATIVO",
      bc: npvNafinPositive ? VIAB_COLORS.dark : VIAB_COLORS.red,
      color: npvNafinPositive ? VIAB_COLORS.dark : VIAB_COLORS.red,
      note: "VPN positivo = proyecto crea valor.",
    },
    {
      l: "VPN @ 10%",
      v: viabFormatSignedK(VIAB_NPV_10),
      s: "Tasa alternativa",
      badge: "REFERENCIA",
      bc: VIAB_COLORS.mid,
      color: VIAB_COLORS.mid,
      note: "Escenario conservador de costo de capital.",
    },
    {
      l: "ROI Año 1",
      v: "128%",
      s: "Revenue/inversión",
      badge: "ALTO",
      bc: VIAB_COLORS.dark,
      color: VIAB_COLORS.dark,
      note: "$98,515 revenue vs $76,960 inversión.",
    },
    {
      l: "ROI Año 2",
      v: "292%",
      s: "Revenue acum./inv.",
      badge: "MUY ALTO",
      bc: VIAB_COLORS.dark,
      color: VIAB_COLORS.dark,
      note: "Revenue acumulado / inversión inicial.",
    },
    {
      l: "Payback",
      v: "~9 meses",
      s: "FCF acumulado > 0",
      badge: "RÁPIDO",
      bc: VIAB_COLORS.mid,
      color: VIAB_COLORS.mid,
      note: "Recuperación antes de cierre de Año 1.",
    },
  ];

  const eficiencia = [
    { l: "Margen bruto/viaje", v: "~83%", s: "Precio $31 vs COGS $5.17", color: VIAB_COLORS.dark, bar: 83, warn: false },
    { l: "Margen EBITDA Año 2", v: "59.0%", s: "$133K / $225K revenue", color: VIAB_COLORS.dark, bar: 59, warn: false },
    { l: "Margen EBITDA Año 3+", v: "61.3%", s: "Plateau desde Año 3", color: VIAB_COLORS.mid, bar: 61.3, warn: false },
    { l: "OPEX ratio Año 1", v: `${VIAB_OPEX_R_Y1}%`, s: "OPEX/Revenue", color: VIAB_COLORS.orange, bar: Number(VIAB_OPEX_R_Y1), warn: true },
    { l: "OPEX ratio Año 2", v: `${VIAB_OPEX_R_Y2}%`, s: "Mejora eficiencia", color: VIAB_COLORS.mid, bar: Number(VIAB_OPEX_R_Y2), warn: true },
    { l: "Rev./sensor Año 2", v: `$${viabFormatInt(Math.round(224919 / 601))}`, s: "USD por sensor activo", color: VIAB_COLORS.dark, bar: 75, warn: false },
  ];

  const waterfall = [
    { l: "Revenue", v: "$224,919", pct: "100%", desc: "10 clientes activos", arrow: "", bg: VIAB_COLORS.light, color: VIAB_COLORS.dark },
    { l: "(-) COGS", v: "-$37,238", pct: "16.6%", desc: "$5.17 × 7,203 viajes", arrow: "↓", bg: "#f5faf7", color: VIAB_COLORS.mid },
    { l: "(-) OPEX", v: "-$55,000", pct: "24.5%", desc: "Personal + infra + admin", arrow: "↓", bg: "#eef7f1", color: VIAB_COLORS.mid },
    { l: "= EBITDA", v: "$132,681", pct: "59.0%", desc: "Margen sobre revenue", arrow: "=", bg: VIAB_COLORS.light, color: VIAB_COLORS.dark },
  ];

  return `
    <section class="viab-section">
      <h3 class="viab-section-title">Indicadores financieros clave</h3>

      <div class="viab-indicator-group">
        <div class="viab-indicator-group-title">Rentabilidad y retorno</div>
        <div class="viab-indicator-grid">
          ${rentabilidad
            .map(
              (k) => `
            <article class="viab-indicator-card">
              <div class="viab-indicator-label">${k.l}</div>
              <div class="viab-indicator-value" style="color:${k.color}">${k.v}</div>
              <div class="viab-indicator-row">
                <span>${k.s}</span>
                <span class="viab-indicator-badge" style="color:${k.bc};border-color:${k.bc}33;background:${k.bc}18">${k.badge}</span>
              </div>
              <div class="viab-indicator-note">${k.note}</div>
            </article>
          `,
            )
            .join("")}
        </div>
      </div>

      <div class="viab-indicator-group">
        <div class="viab-indicator-group-title">Eficiencia operativa</div>
        <div class="viab-indicator-grid">
          ${eficiencia
            .map(
              (k) => `
            <article class="viab-indicator-card">
              <div class="viab-indicator-label">${k.l}</div>
              <div class="viab-indicator-value" style="color:${k.color}">${k.v}</div>
              <div class="viab-indicator-meter">
                <div class="viab-indicator-meter-fill${k.warn ? " is-warn" : ""}" style="width:${Math.min(Number(k.bar) || 0, 100)}%"></div>
              </div>
              <div class="viab-indicator-note">${k.s}</div>
            </article>
          `,
            )
            .join("")}
        </div>
      </div>

      <div class="viab-indicator-group" style="margin-bottom:0">
        <div class="viab-indicator-group-title">Construcción del EBITDA — Año 2 (waterfall)</div>
        <div class="viab-waterfall-grid">
          ${waterfall
            .map(
              (k) => `
            <article class="viab-waterfall-card" style="background:${k.bg}">
              <div class="viab-waterfall-arrow">${k.arrow}</div>
              <div class="viab-indicator-label">${k.l}</div>
              <div class="viab-indicator-value" style="color:${k.color}">${k.v}</div>
              <div class="viab-indicator-row"><span>${k.desc}</span><strong style="color:${k.color}">${k.pct}</strong></div>
            </article>
          `,
            )
            .join("")}
        </div>
      </div>
    </section>
  `;
}

function viabBuildStyles() {
  return `
    <style>
      .viabilidad-panel{padding:1.2rem;border:0;background:transparent;box-shadow:none}
      .viab-root{background:transparent;border:0;border-radius:0;padding:0;font-family:inherit;color:${VIAB_COLORS.textMid};display:flex;flex-direction:column;gap:1.15rem}
      .viab-root *{font-family:inherit}
      .viab-section{min-height:0;background:#fff;border:1px solid #d9e6dd;border-radius:18px;padding:1.2rem;margin-bottom:0}
      .viab-banner{background:${VIAB_COLORS.dark};color:#fff;border:1px solid ${VIAB_COLORS.dark};border-radius:18px;min-height:0;padding:1.2rem;margin:0;width:auto}
      .viab-banner-title{margin:0 0 0.85rem;font-size:clamp(1.15rem, 2.3vw, 1.55rem);line-height:1.1;font-weight:800;color:#ffffff}
      .viab-banner-grid{width:100%;max-width:100%;margin:0 auto;padding:0;display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}
      .viab-banner-card{border:1px solid rgba(255,255,255,0.26);border-radius:12px;background:rgba(255,255,255,0.08);padding:0.72rem 0.68rem;text-align:center}
      .viab-banner-label{font-size:0.7rem;color:#d9f0df;letter-spacing:0.01em;margin-bottom:0.22rem}
      .viab-banner-value{font-size:clamp(0.9rem, 1.7vw, 1.2rem);line-height:1.15;font-weight:800}
      .viab-banner-sub{font-size:0.7rem;color:#d9f0df;letter-spacing:0.01em;margin-top:0.22rem}
      .viab-kpi-grid{margin-bottom:16px;display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px;width:100%}
      .viab-kpi-card{background:#fff;border:1px solid #d9e6dd;border-radius:12px;padding:0.72rem 0.7rem;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center}
      .viab-kpi-card.is-accent{background:${VIAB_COLORS.light}}
      .viab-kpi-card.kpi-recuperacion{display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center}
      .viab-kpi-title{display:block;font-size:0.92rem;font-weight:800;color:#1a6b3a;margin:0 0 0.28rem;line-height:1.18;text-align:center}
      .viab-kpi-title .viab-unit{display:block;width:max-content;margin:0.22rem auto 0}
      .viab-unit{display:inline-block;background:#f1f8f3;color:#1a6b3a;border:1px solid #c9dece;border-radius:999px;padding:0.12rem 0.46rem;font-size:0.7rem;font-weight:700}
      .viab-kpi-value{font-size:clamp(0.9rem, 1.7vw, 1.2rem);font-weight:400;line-height:1.15;color:#111111}
      .viab-kpi-sub{font-size:0.79rem;color:#355264;margin-top:0.18rem}
      .viab-section-title{margin:0 0 0.72rem;color:#1a6b3a;font-size:1.26rem;font-weight:800;line-height:1.2}
      .viab-two-col{display:grid;grid-template-columns:1fr 1fr;gap:24px}
      .viab-three-col{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px}
      .viab-scenario-controls{display:flex;gap:6px;flex-wrap:wrap}
      .viab-scenario-btn{cursor:pointer;border:1.5px solid ${VIAB_COLORS.grayB};background:#fff;border-radius:999px;padding:4px 13px;font-size:11px;color:${VIAB_COLORS.textMid};font-weight:600}
      .viab-scenario-btn:hover{border-color:${VIAB_COLORS.dark};color:${VIAB_COLORS.dark}}
      .viab-scenario-btn.is-active{border-color:${VIAB_COLORS.dark};background:${VIAB_COLORS.dark};color:#fff}
      .viab-table-wrap{overflow-x:auto}
      .viab-table{border-collapse:collapse;width:100%}
      .viab-table th{background:${VIAB_COLORS.grayL};color:${VIAB_COLORS.textDim};font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;padding:7px 9px;text-align:right;border-bottom:1px solid ${VIAB_COLORS.grayB}}
      .viab-table th:first-child,.viab-table td:first-child{text-align:left}
      .viab-table td{font-size:11px;padding:6px 9px;text-align:right;border-bottom:1px solid ${VIAB_COLORS.grayB};color:${VIAB_COLORS.textMid}}
      .viab-table tr:last-child td{border-bottom:none}
      .viab-highlight-row td{color:${VIAB_COLORS.dark};font-weight:700;background:${VIAB_COLORS.light}!important}
      .viab-neg{color:#dc2626!important}
      .viab-note{margin-top:10px;padding:8px 12px;background:${VIAB_COLORS.light};border-radius:6px;font-size:11px;color:${VIAB_COLORS.dark}}
      .viab-membership-bar-row{margin-bottom:10px}
      .viab-membership-bar-head{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:3px;font-size:11px}
      .viab-membership-track{height:5px;background:${VIAB_COLORS.grayB};border-radius:3px}
      .viab-membership-fill{height:100%;border-radius:3px;opacity:.78}
      .viab-membership-caption{font-size:10px;color:${VIAB_COLORS.textDim};margin-top:2px}
      .viab-chip{display:inline-block;border:1px solid;border-radius:999px;padding:0 6px;font-size:9px;font-weight:700;margin-left:6px}
      .viab-chart-row{display:grid;grid-template-columns:1fr 1fr;gap:20px}
      .viab-chart-title{font-size:11px;color:${VIAB_COLORS.textDim};font-weight:600;margin:0 0 8px}
      .viab-chart-box{height:220px}
      .viab-legend{display:flex;gap:14px;margin-top:6px;font-size:11px}
      .viab-summary-grid{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px;margin-top:16px;padding-top:14px;border-top:1px solid ${VIAB_COLORS.grayB}}
      .viab-summary-item{text-align:center}
      .viab-summary-label{font-size:9px;color:${VIAB_COLORS.textDim};text-transform:uppercase;letter-spacing:.05em;margin-bottom:3px}
      .viab-summary-value{font-size:14px;font-weight:800;color:${VIAB_COLORS.dark}}
      .viab-indicator-group{margin-bottom:16px}
      .viab-indicator-group-title{font-size:11px;font-weight:700;color:${VIAB_COLORS.textDim};text-transform:uppercase;letter-spacing:.08em;margin:0 0 10px}
      .viab-indicator-grid{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px}
      .viab-indicator-card{background:${VIAB_COLORS.grayL};border:1px solid ${VIAB_COLORS.grayB};border-radius:8px;padding:11px 13px}
      .viab-indicator-label{font-size:9px;color:${VIAB_COLORS.textDim};margin-bottom:5px;font-weight:600;text-transform:uppercase;letter-spacing:.05em}
      .viab-indicator-value{font-size:19px;font-weight:800;margin-bottom:5px}
      .viab-indicator-row{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:5px;font-size:9px;color:${VIAB_COLORS.textDim}}
      .viab-indicator-badge{display:inline-block;font-size:8px;padding:1px 6px;border-radius:999px;border:1px solid;font-weight:700}
      .viab-indicator-note{font-size:9px;color:${VIAB_COLORS.textDim};line-height:1.4}
      .viab-indicator-meter{height:4px;background:${VIAB_COLORS.grayB};border-radius:2px;margin-bottom:5px}
      .viab-indicator-meter-fill{height:100%;border-radius:2px;background:${VIAB_COLORS.dark};opacity:.75}
      .viab-indicator-meter-fill.is-warn{background:${VIAB_COLORS.orange}}
      .viab-waterfall-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}
      .viab-waterfall-card{border:1px solid ${VIAB_COLORS.grayB};border-radius:8px;padding:12px 14px;position:relative}
      .viab-waterfall-arrow{position:absolute;left:-9px;top:50%;transform:translateY(-50%);font-size:16px;font-weight:700;color:${VIAB_COLORS.textDim}}
      .viab-client-card{border:1px solid ${VIAB_COLORS.grayB};border-left:3px solid ${VIAB_COLORS.orange};border-radius:8px;padding:11px 14px;background:#fff;margin-bottom:8px}
      .viab-client-card.is-muted{opacity:.86}
      .viab-client-head{display:flex;justify-content:space-between;align-items:flex-start;gap:8px}
      .viab-client-quarter{font-size:12px;font-weight:700;color:${VIAB_COLORS.dark}}
      .viab-tier-badge{display:inline-block;background:${VIAB_COLORS.light};color:${VIAB_COLORS.dark};border:1px solid ${VIAB_COLORS.lightB};border-radius:999px;padding:1px 7px;font-size:10px;font-weight:600;margin-left:7px}
      .viab-risk-badge{display:inline-block;border-radius:999px;padding:1px 7px;font-size:10px;font-weight:700;margin-left:6px;border:1px solid}
      .viab-risk-critico{color:${VIAB_COLORS.red};background:${VIAB_COLORS.red}18;border-color:${VIAB_COLORS.red}44}
      .viab-risk-alto{color:${VIAB_COLORS.orange};background:${VIAB_COLORS.orange}18;border-color:${VIAB_COLORS.orange}44}
      .viab-risk-moderado{color:${VIAB_COLORS.yellow};background:${VIAB_COLORS.yellow}18;border-color:${VIAB_COLORS.yellow}44}
      .viab-client-meta{font-size:11px;color:${VIAB_COLORS.textDim};margin-top:4px}
      .viab-client-foot{display:flex;justify-content:space-between;align-items:center;margin-top:5px;font-size:11px;color:${VIAB_COLORS.textDim}}
      .viab-total-row{display:flex;justify-content:space-between;padding:8px 12px;background:${VIAB_COLORS.grayL};border-radius:6px;margin:2px 0 14px;font-size:12px}
      .viab-total-row strong{font-size:13px;color:${VIAB_COLORS.dark}}
      .viab-total-row.is-accent{background:${VIAB_COLORS.light};margin-top:6px}
      .viab-progress-row{margin-bottom:9px}
      .viab-progress-head{display:flex;justify-content:space-between;gap:8px;margin-bottom:3px;font-size:11px}
      .viab-progress-head .is-warn{color:${VIAB_COLORS.orange};font-weight:700}
      .viab-progress-head small{color:${VIAB_COLORS.textDim};font-weight:400;font-size:10px}
      .viab-progress-track{height:4px;background:${VIAB_COLORS.grayB};border-radius:2px}
      .viab-progress-fill{height:100%;background:${VIAB_COLORS.dark};border-radius:2px}
      .viab-progress-fill.is-warn{background:${VIAB_COLORS.orange}}
      .viab-mini-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;margin-bottom:14px}
      .viab-mini-card{background:${VIAB_COLORS.light};border-radius:7px;padding:9px 12px}
      .viab-mini-label{font-size:10px;color:${VIAB_COLORS.textDim};margin-bottom:3px}
      .viab-mini-value{font-size:17px;font-weight:800;color:${VIAB_COLORS.dark}}
      .viab-warn{margin-top:8px;padding:7px 10px;background:#fff3e0;border-radius:6px;border:1px solid ${VIAB_COLORS.orange}44;font-size:11px;color:#7c3c00}
      .viab-positive{color:${VIAB_COLORS.mid}!important;font-weight:600}
      .viab-strong{color:${VIAB_COLORS.dark}!important;font-weight:700}
      .viab-scenario-card{border:1px solid ${VIAB_COLORS.grayB};border-top:3px solid ${VIAB_COLORS.dark};border-radius:8px;padding:14px 16px}
      .viab-scenario-name{font-weight:700;font-size:13px;margin-bottom:10px}
      .viab-scenario-kv{display:flex;justify-content:space-between;padding-bottom:6px;margin-bottom:6px;border-bottom:1px solid ${VIAB_COLORS.grayB};font-size:11px;color:${VIAB_COLORS.textDim}}
      .viab-scenario-kv:last-of-type{border-bottom:0;margin-bottom:0}
      .viab-scenario-kv strong{font-size:12px}
      .viab-scenario-desc{font-size:10px;color:${VIAB_COLORS.textDim};margin:6px 0 0;font-style:italic}
      .viab-footer{border-top:1px solid ${VIAB_COLORS.grayB};padding-top:14px;display:flex;justify-content:space-between;gap:12px;align-items:center;font-size:11px;color:${VIAB_COLORS.textDim};flex-wrap:wrap}
      @media (max-width:1200px){
        .viab-kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
        .viab-summary-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
        .viab-kpi-card.kpi-recuperacion{grid-column:1 / -1}
        .viab-indicator-grid{grid-template-columns:repeat(3,minmax(0,1fr))}
        .viab-waterfall-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
      }
      @media (max-width:900px){
        .viab-kpi-grid{grid-template-columns:1fr}
        .viab-three-col,.viab-chart-row,.viab-two-col,.viab-indicator-grid,.viab-waterfall-grid{grid-template-columns:1fr}
        .viab-waterfall-arrow{display:none}
      }
      @media (max-width:720px){
        .viab-banner-grid{grid-template-columns:1fr}
      }
    </style>
  `;
}

function renderViabilidadTab() {
  const container = document.getElementById("viabilidadContainer");
  if (!container) return;
  const scenario = ["Base", "Conservador", "Optimista"].includes(state.viabilidadScenario)
    ? state.viabilidadScenario
    : "Base";
  const data = getViabilidadScenarioData();
  const scenarioButtons = viabScenarioButtonsHtml(scenario);

  container.innerHTML = `
    ${viabBuildStyles()}
    <div class="viab-root">
      <section class="viab-banner">
        <h2 class="viab-banner-title">Viabilidad Financiera</h2>
        <div class="viab-banner-grid kpi-cards-wrapper">
          <div class="viab-banner-card"><div class="viab-banner-label">Inversión Año 1</div><div class="viab-banner-value">$76,960 USD</div><div class="viab-banner-sub">6 rubros</div></div>
          <div class="viab-banner-card"><div class="viab-banner-label">Revenue Año 1</div><div class="viab-banner-value">$98,515 USD</div><div class="viab-banner-sub">5 clientes · 3,069 viajes</div></div>
          <div class="viab-banner-card"><div class="viab-banner-label">Revenue Año 2</div><div class="viab-banner-value">$224,919 USD</div><div class="viab-banner-sub">Renovaciones + 5 nuevos</div></div>
          <div class="viab-banner-card"><div class="viab-banner-label">Payback estimado</div><div class="viab-banner-value">~9 meses</div><div class="viab-banner-sub">FCF acumulado > 0</div></div>
        </div>
      </section>

      <section class="viab-section">
        <div class="viab-kpi-grid kpi-cards-wrapper" style="margin-bottom:0">
          <article class="viab-kpi-card"><h4 class="viab-kpi-title">COGS por viaje <span class="viab-unit">USD</span></h4><div class="viab-kpi-value">$5.17</div><div class="viab-kpi-sub">vs spot $27.50</div></article>
          <article class="viab-kpi-card"><h4 class="viab-kpi-title">Precio prom. ponderado <span class="viab-unit">USD/viaje</span></h4><div class="viab-kpi-value">$31.22</div><div class="viab-kpi-sub">Año 2 · 7,203 viajes</div></article>
          <article class="viab-kpi-card is-accent"><h4 class="viab-kpi-title">Margen bruto / viaje <span class="viab-unit">%</span></h4><div class="viab-kpi-value">~83%</div><div class="viab-kpi-sub">Antes de OPEX</div></article>
          <article class="viab-kpi-card"><h4 class="viab-kpi-title">Sensores Año 2 <span class="viab-unit">unidades</span></h4><div class="viab-kpi-value">601 u.</div><div class="viab-kpi-sub">256 Y1 + 345 Y2</div></article>
          <article class="viab-kpi-card is-accent kpi-recuperacion"><h4 class="viab-kpi-title">Recuperación fiscal <span class="viab-unit">LISR</span></h4><div class="viab-kpi-value">51–65%</div><div class="viab-kpi-sub">Art.189 + Art.25</div></article>
        </div>
      </section>

      ${viabRenderIndicadoresSection()}

      <section class="viab-section">
        <h3 class="viab-section-title">Desglose de membresías — Tiers de precio por volumen</h3>
        <div class="viab-two-col">
          <div>
            <div class="viab-chart-title">Estructura de tiers <span class="viab-unit">USD / viaje</span></div>
            <div class="viab-table-wrap">
              <table class="viab-table">
                <thead>
                  <tr><th>Tier</th><th>Rango DUA/mes</th><th>Precio/viaje</th><th>COGS/viaje</th><th>Margen bruto</th></tr>
                </thead>
                <tbody>
                  <tr><td style="font-weight:700;color:${VIAB_COLORS.textDim}">Sin Tier</td><td>1–49 DUA/mes</td><td style="font-weight:700;color:${VIAB_COLORS.textDim}">$35</td><td>$5.17</td><td class="viab-positive">85.2%</td></tr>
                  <tr><td style="font-weight:700;color:${VIAB_COLORS.mid}">Explorador</td><td>50–149 DUA/mes</td><td style="font-weight:700;color:${VIAB_COLORS.mid}">$32</td><td>$5.17</td><td class="viab-positive">83.8%</td></tr>
                  <tr><td style="font-weight:700;color:${VIAB_COLORS.mid}">Socio</td><td>150–299 DUA/mes</td><td style="font-weight:700;color:${VIAB_COLORS.mid}">$30</td><td>$5.17</td><td class="viab-positive">82.8%</td></tr>
                  <tr><td style="font-weight:700;color:${VIAB_COLORS.dark}">Estrategico</td><td>300–499 DUA/mes</td><td style="font-weight:700;color:${VIAB_COLORS.dark}">$28</td><td>$5.17</td><td class="viab-positive">81.5%</td></tr>
                  <tr><td style="font-weight:700;color:${VIAB_COLORS.dark}">Ancla</td><td>500+ DUA/mes</td><td style="font-weight:700;color:${VIAB_COLORS.dark}">$25</td><td>$5.17</td><td class="viab-positive">79.3%</td></tr>
                </tbody>
              </table>
            </div>
            <p class="viab-note">Try-out: 1–3 meses renovable · $33–$35/viaje · Sin compromiso de volumen · Cláusula comodato $150/sensor perdido</p>
          </div>
          <div>
            <div class="viab-chart-title">Revenue por cliente Año 1 — tier asignado</div>
            ${viabRenderMembershipBars()}
            <div style="border-top:1px solid ${VIAB_COLORS.grayB};padding-top:8px;margin-top:4px;display:flex;justify-content:space-between;font-size:12px"><span style="color:${VIAB_COLORS.textDim}">Precio ponderado Año 1</span><strong style="color:${VIAB_COLORS.dark}">$32.10 / viaje</strong></div>
            <div style="display:flex;justify-content:space-between;font-size:12px;margin-top:4px"><span style="color:${VIAB_COLORS.textDim}">Precio ponderado Año 2</span><strong style="color:${VIAB_COLORS.dark}">$31.22 / viaje</strong></div>
          </div>
        </div>
      </section>

      <section class="viab-section">
        <div style="display:flex;justify-content:space-between;gap:10px;align-items:center;margin-bottom:16px;flex-wrap:wrap">
          <h3 class="viab-section-title" style="margin:0">Revenue · EBITDA · FCF Acumulado <span class="viab-unit">USD</span></h3>
          <div class="viab-scenario-controls">${scenarioButtons}</div>
        </div>
        <div class="viab-chart-row">
          <div>
            <p class="viab-chart-title">Revenue y EBITDA — proyección 10 años</p>
            <div id="viabChartRevenue" class="viab-chart-box"></div>
            <div class="viab-legend">
              <span style="color:${VIAB_COLORS.dark};font-weight:600">— Revenue</span>
              <span style="color:${VIAB_COLORS.mid};font-weight:600">— EBITDA</span>
              <span style="color:${VIAB_COLORS.dark};font-weight:600">--- Util. Neta</span>
            </div>
          </div>
          <div>
            <p class="viab-chart-title">FCF Acumulado post-inversión — 3 escenarios</p>
            <div id="viabChartFcf" class="viab-chart-box"></div>
            <div class="viab-legend">
              <span style="color:${VIAB_COLORS.dark};font-weight:600">— Base ($958K)</span>
              <span style="color:${VIAB_COLORS.orange};font-weight:600">--- Cons. ($430K)</span>
              <span style="color:${VIAB_COLORS.mid};font-weight:600">--- Opt. ($1.5M)</span>
            </div>
          </div>
        </div>
        <div class="viab-summary-grid">
          <div class="viab-summary-item"><div class="viab-summary-label">Revenue Año 1</div><div class="viab-summary-value">$98,515</div></div>
          <div class="viab-summary-item"><div class="viab-summary-label">Revenue Año 2</div><div class="viab-summary-value">$224,919</div></div>
          <div class="viab-summary-item"><div class="viab-summary-label">EBITDA Año 2</div><div class="viab-summary-value">$132,681</div></div>
          <div class="viab-summary-item"><div class="viab-summary-label">Margen EBITDA</div><div class="viab-summary-value">59.0%</div></div>
          <div class="viab-summary-item"><div class="viab-summary-label">FCF Base Año 10</div><div class="viab-summary-value">$958K acum</div></div>
        </div>
      </section>

      <section class="viab-two-col" style="margin-bottom:0">
        <article class="viab-section" style="margin:0">
          <h3 class="viab-section-title">Crédito NAFIN Eco Crédito <span class="viab-unit">Revolvente</span></h3>
          <div class="viab-mini-grid">
            <div class="viab-mini-card"><div class="viab-mini-label">Línea total</div><div class="viab-mini-value">$76,960</div></div>
            <div class="viab-mini-card"><div class="viab-mini-label">Tasa anual fija</div><div class="viab-mini-value">16.0%</div></div>
            <div class="viab-mini-card"><div class="viab-mini-label">Plazo</div><div class="viab-mini-value">36 meses</div></div>
          </div>
          <div style="font-size:12px;font-weight:700;color:${VIAB_COLORS.dark};margin-bottom:8px">Tramos de disposición</div>
          ${viabRenderDrawdownBars()}
          <div style="font-size:12px;font-weight:700;color:${VIAB_COLORS.dark};margin:12px 0 8px">Saldo deudor — primer año</div>
          <div id="viabChartAmort" style="height:110px"></div>
        </article>
        <article class="viab-section" style="margin:0">
          <h3 class="viab-section-title">Desglose de inversión Año 1 <span class="viab-unit">$76,960 USD</span></h3>
          ${viabRenderInvestmentBars()}
          <p class="viab-warn">⚑ Legal & Fiscal incluye resolución SAT Art.189 ($10K) — requerida antes del pitch a CFOs</p>
        </article>
      </section>

      <section class="viab-section">
        <h3 class="viab-section-title">Beneficios fiscales por cliente <span class="viab-unit">LISR Art. 189 + Art. 25</span></h3>
        <div class="viab-two-col">
          <div>
            <div style="background:${VIAB_COLORS.light};border-radius:8px;padding:12px 14px;margin-bottom:14px;border-left:3px solid ${VIAB_COLORS.dark}">
              <div style="font-weight:700;color:${VIAB_COLORS.dark};margin-bottom:6px;font-size:13px">Mecánica de recuperación fiscal</div>
              <div style="font-size:12px;color:${VIAB_COLORS.textMid};line-height:1.7">
                <strong>Art. 189 LISR:</strong> Crédito directo → 30% del costo descuenta peso a peso del ISR<br/>
                <strong>Art. 25 LISR:</strong> 70% del gasto deducible × 30% ISR = 21% adicional<br/>
                <strong style="color:${VIAB_COLORS.dark}">Recuperación estándar: 51% · Clientes IMMEX: hasta 65%</strong>
              </div>
            </div>
            <div id="viabChartFiscal" style="height:160px"></div>
            <div class="viab-legend"><span style="color:${VIAB_COLORS.dark};font-weight:600">■ Art.189 (30%)</span><span style="color:${VIAB_COLORS.mid};font-weight:600">■ Art.25 (21%)</span><span style="color:${VIAB_COLORS.textDim}">= 51% recuperación</span></div>
          </div>
          <div>
            <div style="font-size:12px;font-weight:700;color:${VIAB_COLORS.dark};margin-bottom:8px">Costo neto vs mercado spot ($27.50/viaje)</div>
            <div class="viab-table-wrap">
              <table class="viab-table">
                <thead><tr><th>Cliente</th><th>Membresía</th><th>Recup. 51%</th><th>Costo neto</th><th>Ahorro anual</th></tr></thead>
                <tbody>${viabRenderCostoNetoRows()}</tbody>
              </table>
            </div>
          </div>
        </div>
      </section>

      <section class="viab-section">
        <h3 class="viab-section-title">Análisis de escenarios — Métricas clave Año 10 (2035)</h3>
        <div class="viab-three-col" style="margin-bottom:16px">${viabRenderScenariosCards()}</div>
        <div id="viabChartScenario" style="height:130px"></div>
      </section>

      <footer class="viab-footer">
        <span>© CL Circular · Modelo financiero México 2026–2035 · Uso interno</span>
        <span>COGS $5.17/viaje · NAFIN 16% · ISR 30% · Renovación sensores Y4 y Y8 · Pérdida 5%/año</span>
      </footer>
    </div>
  `;

  container.querySelectorAll(".viab-scenario-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const next = btn.getAttribute("data-viab-scenario");
      if (!next || next === state.viabilidadScenario) return;
      state.viabilidadScenario = next;
      renderViabilidadTab();
      requestAnimationFrame(() => resizeAllCharts());
    });
  });

  renderViabilidadCharts(data);
}

function renderViabilidadCharts(data) {
  if (typeof Plotly === "undefined") {
    ["viabChartRevenue", "viabChartFcf", "viabChartAmort", "viabChartFiscal", "viabChartScenario"].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.innerHTML = `<p style="font-size:12px;color:${VIAB_COLORS.textDim};margin:0">No se pudo cargar Plotly.</p>`;
    });
    return;
  }

  const years = data.map((d) => d.yr);
  Plotly.react(
    "viabChartRevenue",
    [
      {
        x: years,
        y: data.map((d) => d.rev),
        type: "scatter",
        mode: "lines",
        name: "Revenue",
        line: { color: VIAB_COLORS.dark, width: 2.5 },
        fill: "tozeroy",
        fillcolor: "rgba(26,107,58,0.14)",
      },
      {
        x: years,
        y: data.map((d) => d.ebitda),
        type: "scatter",
        mode: "lines",
        name: "EBITDA",
        line: { color: VIAB_COLORS.mid, width: 2 },
        fill: "tozeroy",
        fillcolor: "rgba(46,143,79,0.10)",
      },
      {
        x: years,
        y: data.map((d) => d.ni),
        type: "scatter",
        mode: "lines",
        name: "Util. Neta",
        line: { color: VIAB_COLORS.dark, width: 1.5, dash: "dash" },
      },
    ],
    {
      margin: { t: 6, r: 8, b: 30, l: 42 },
      paper_bgcolor: "#ffffff",
      plot_bgcolor: "#ffffff",
      showlegend: false,
      xaxis: { tickfont: { size: 9, color: VIAB_COLORS.textDim }, showgrid: false, zeroline: false },
      yaxis: { tickfont: { size: 9, color: VIAB_COLORS.textDim }, gridcolor: VIAB_COLORS.grayB, tickprefix: "$", tickformat: "~s", zeroline: false },
    },
    { responsive: true, displaylogo: false, modeBarButtonsToRemove: ["select2d", "lasso2d"] },
  );

  Plotly.react(
    "viabChartFcf",
    [
      {
        x: VIAB_CUM_CHART.map((d) => d.yr),
        y: VIAB_CUM_CHART.map((d) => d.Base),
        type: "scatter",
        mode: "lines+markers",
        name: "Base",
        line: { color: VIAB_COLORS.dark, width: 2.5 },
        marker: { size: 6 },
      },
      {
        x: VIAB_CUM_CHART.map((d) => d.yr),
        y: VIAB_CUM_CHART.map((d) => d.Conservador),
        type: "scatter",
        mode: "lines",
        name: "Conservador",
        line: { color: VIAB_COLORS.orange, width: 1.6, dash: "dash" },
      },
      {
        x: VIAB_CUM_CHART.map((d) => d.yr),
        y: VIAB_CUM_CHART.map((d) => d.Optimista),
        type: "scatter",
        mode: "lines",
        name: "Optimista",
        line: { color: VIAB_COLORS.mid, width: 1.6, dash: "dash" },
      },
    ],
    {
      margin: { t: 6, r: 8, b: 30, l: 42 },
      paper_bgcolor: "#ffffff",
      plot_bgcolor: "#ffffff",
      showlegend: false,
      xaxis: { tickfont: { size: 9, color: VIAB_COLORS.textDim }, showgrid: false, zeroline: false },
      yaxis: { tickfont: { size: 9, color: VIAB_COLORS.textDim }, gridcolor: VIAB_COLORS.grayB, ticksuffix: "K", zeroline: false },
      shapes: [
        {
          type: "line",
          x0: 0,
          x1: 1,
          xref: "paper",
          y0: 0,
          y1: 0,
          yref: "y",
          line: { color: VIAB_COLORS.grayB, dash: "dot", width: 1.3 },
        },
      ],
      annotations: [
        {
          text: "Break-even",
          xref: "paper",
          x: 0.02,
          y: 0,
          yref: "y",
          showarrow: false,
          font: { size: 9, color: VIAB_COLORS.dark },
          yshift: 10,
        },
      ],
    },
    { responsive: true, displaylogo: false, modeBarButtonsToRemove: ["select2d", "lasso2d"] },
  );

  Plotly.react(
    "viabChartAmort",
    [
      {
        x: VIAB_AMORT.map((d) => d.mes),
        y: VIAB_AMORT.map((d) => d.capital),
        type: "bar",
        name: "Capital",
        marker: { color: VIAB_COLORS.dark, opacity: 0.86 },
      },
      {
        x: VIAB_AMORT.map((d) => d.mes),
        y: VIAB_AMORT.map((d) => d.interes),
        type: "bar",
        name: "Interés",
        marker: { color: VIAB_COLORS.orange, opacity: 0.72 },
      },
    ],
    {
      barmode: "stack",
      margin: { t: 4, r: 6, b: 20, l: 36 },
      paper_bgcolor: "#ffffff",
      plot_bgcolor: "#ffffff",
      showlegend: false,
      xaxis: { tickfont: { size: 8, color: VIAB_COLORS.textDim }, showgrid: false, zeroline: false },
      yaxis: { tickfont: { size: 8, color: VIAB_COLORS.textDim }, ticksuffix: "K", tickformat: "~s", gridcolor: VIAB_COLORS.grayB, zeroline: false },
    },
    { responsive: true, displaylogo: false, modeBarButtonsToRemove: ["select2d", "lasso2d"] },
  );

  Plotly.react(
    "viabChartFiscal",
    [
      {
        y: VIAB_FISCAL.map((d) => d.name),
        x: VIAB_FISCAL.map((d) => d.art189),
        type: "bar",
        orientation: "h",
        name: "Art.189 (30%)",
        marker: { color: VIAB_COLORS.dark, opacity: 0.86 },
      },
      {
        y: VIAB_FISCAL.map((d) => d.name),
        x: VIAB_FISCAL.map((d) => d.art25),
        type: "bar",
        orientation: "h",
        name: "Art.25 (21%)",
        marker: { color: VIAB_COLORS.mid, opacity: 0.72 },
      },
    ],
    {
      barmode: "stack",
      margin: { t: 2, r: 8, b: 20, l: 50 },
      paper_bgcolor: "#ffffff",
      plot_bgcolor: "#ffffff",
      showlegend: false,
      xaxis: { tickfont: { size: 8, color: VIAB_COLORS.textDim }, ticksuffix: "K", tickformat: "~s", gridcolor: VIAB_COLORS.grayB, zeroline: false },
      yaxis: { tickfont: { size: 9, color: VIAB_COLORS.textMid }, showgrid: false, zeroline: false },
    },
    { responsive: true, displaylogo: false, modeBarButtonsToRemove: ["select2d", "lasso2d"] },
  );

  Plotly.react(
    "viabChartScenario",
    [
      {
        x: ["Revenue Y10", "EBITDA Y10", "Util.Neta Y10", "FCF Acum."],
        y: [285, 175, 73, 958],
        type: "bar",
        name: "Base",
        marker: { color: VIAB_COLORS.dark, opacity: 0.85 },
      },
      {
        x: ["Revenue Y10", "EBITDA Y10", "Util.Neta Y10", "FCF Acum."],
        y: [198, 96, 38, 430],
        type: "bar",
        name: "Conservador",
        marker: { color: VIAB_COLORS.orange, opacity: 0.72 },
      },
      {
        x: ["Revenue Y10", "EBITDA Y10", "Util.Neta Y10", "FCF Acum."],
        y: [398, 219, 95, 1500],
        type: "bar",
        name: "Optimista",
        marker: { color: VIAB_COLORS.mid, opacity: 0.72 },
      },
    ],
    {
      barmode: "group",
      margin: { t: 4, r: 8, b: 24, l: 42 },
      paper_bgcolor: "#ffffff",
      plot_bgcolor: "#ffffff",
      showlegend: false,
      xaxis: { tickfont: { size: 9, color: VIAB_COLORS.textDim }, showgrid: false, zeroline: false },
      yaxis: { tickfont: { size: 8, color: VIAB_COLORS.textDim }, ticksuffix: "K", gridcolor: VIAB_COLORS.grayB, zeroline: false },
    },
    { responsive: true, displaylogo: false, modeBarButtonsToRemove: ["select2d", "lasso2d"] },
  );
}

function initPropuestaCoverageMap() {
  const mapEl = document.getElementById("propuestaCoverageMap");
  if (!mapEl || typeof L === "undefined") return;

  if (state.propuestaCoverageMap) {
    state.propuestaCoverageMap.remove();
    state.propuestaCoverageMap = null;
  }

  const map = L.map(mapEl, { zoomControl: false }).setView([27.5, -111.5], 4.8);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);

  const points = getPropuestaCoveragePoints();
  const bounds = [];

  points.forEach((point) => {
    const marker = L.marker([point.lat, point.lng], {
      icon: buildProspectProfileIcon(point.perfil),
    })
      .addTo(map)
      .bindPopup(
        `
        <div class="prospect-popup">
          <strong>${escapeHtml(point.nombre)}</strong><br/>
          <span><strong>Perfil:</strong> ${escapeHtml(point.perfil)}</span><br/>
          <span><strong>Envíos/año:</strong> ${escapeHtml(point.envios)}</span><br/>
          <span><strong>Trimestre:</strong> ${escapeHtml(point.trimestre)}</span>
        </div>
      `,
      );
    bounds.push(marker.getLatLng());
  });

  if (bounds.length) {
    map.fitBounds(bounds, { padding: [42, 42], maxZoom: 6 });
  }

  addProspectMapLegend(map);

  state.propuestaCoverageMap = map;
}

function getPropuestaCoveragePoints() {
  const clusterRows = getClusteredProspectData().rows;
  const clusterByKey = new Map(clusterRows.map((row) => [normalizeCompanyKey(row.empresa), row]));
  const basePoints = [
    { nombre: "Grupo Pinsa", lat: 23.2494, lng: -106.4111, perfil: "Ancla", enviosFallback: "513 envíos/año", trimestre: "Q2 2026" },
    { nombre: "GAM", lat: 23.2194, lng: -106.4411, perfil: "Ancla", enviosFallback: "720 envíos/año", trimestre: "Q2 2026" },
    {
      nombre: "Baja Aqua-Farms",
      lat: 31.88,
      lng: -116.59,
      perfil: "Estratégico",
      enviosFallback: "1,060 envíos/año",
      trimestre: "Q3 2026",
    },
    {
      nombre: "Pacífico Aquaculture",
      lat: 31.84,
      lng: -116.62,
      perfil: "Estratégico",
      enviosFallback: "296 envíos/año",
      trimestre: "Q3 2026",
    },
    {
      nombre: "Baja Shellfish Farms",
      lat: 31.86,
      lng: -116.65,
      perfil: "Estratégico",
      enviosFallback: "480 envíos/año",
      trimestre: "Q4 2026",
    },
  ];
  return basePoints.map((point) => {
    const empresa = findEmpresaByProspectName(point.nombre);
    const clusterRow = clusterByKey.get(normalizeCompanyKey(point.nombre));
    const override = getEstrategiaProspectOverride(empresa?.empresa || point.nombre);
    const envios = formatEnviosLabel(empresa?.viajesAnuales2026 || empresa?.volumenEstimado || point.enviosFallback || "");
    return {
      ...point,
      perfil: override?.profile || clusterRow?.profile || point.perfil,
      trimestre: override?.quarter || clusterRow?.quarter || point.trimestre,
      envios,
    };
  });
}

function buildProspectProfileIcon(perfil = "Estratégico") {
  const normalized = normalizeHeader(perfil).includes("ancla") ? "Ancla" : "Estratégico";
  const profileClass = normalized === "Ancla" ? "prospect-marker-ancla" : "prospect-marker-estrategico";
  return L.divIcon({
    className: "prospect-marker-wrap",
    html: `<div class="prospect-marker prospect-marker-profile ${profileClass}">${escapeHtml(normalized)}</div>`,
    iconSize: [88, 28],
    iconAnchor: [44, 14],
    popupAnchor: [0, -12],
  });
}

function syncPropuestaProspectsLocations() {
  const cards = document.querySelectorAll("#tab-propuesta .propuesta-prospect-card");
  if (!cards.length) return;
  const anclaList = document.querySelector("#tab-propuesta .propuesta-prospect-list-ancla");
  const estrategicoList = document.querySelector("#tab-propuesta .propuesta-prospect-list-estrategico");
  const clusterRows = getClusteredProspectData().rows;
  const clusterByKey = new Map(clusterRows.map((row) => [normalizeCompanyKey(row.empresa), row]));

  cards.forEach((card) => {
    const nameEl = card.querySelector(".propuesta-prospect-head strong");
    const metaEl = card.querySelector(".propuesta-prospect-meta");
    const badgeEl = card.querySelector(".propuesta-profile-badge");
    const quarterEl = card.querySelector(".propuesta-prospect-quarter");
    if (!nameEl || !metaEl) return;

    const empresa = findEmpresaByProspectName(nameEl.textContent || "");
    if (!empresa) return;
    const clusterRow = clusterByKey.get(normalizeCompanyKey(empresa.empresa || nameEl.textContent || ""));
    const override = getEstrategiaProspectOverride(empresa.empresa || nameEl.textContent || "");
    const profile = override?.profile || clusterRow?.profile || "Estratégico";
    const quarter = override?.quarter || clusterRow?.quarter || quarterEl?.textContent || "";

    const ubicacion = (compactRouteText(empresa.ubicacion || "") || NO_INFO).replace(/\.\s*$/, "");
    const duaMensualLabel = formatDuaMesLabelFromEmpresa(empresa);
    const parts = String(metaEl.textContent || "")
      .split("·")
      .map((part) => part.trim())
      .filter(Boolean);
    const prefix = duaMensualLabel && duaMensualLabel !== NO_INFO ? duaMensualLabel : parts[0] || "";
    metaEl.textContent = prefix ? `${prefix} · ${ubicacion}` : ubicacion;

    if (badgeEl) badgeEl.textContent = profile;
    if (quarterEl && quarter) quarterEl.textContent = quarter;

    card.classList.remove("propuesta-q1", "propuesta-q2", "propuesta-q3");
    card.classList.add(profile === "Ancla" ? "propuesta-q2" : "propuesta-q3");

    if (anclaList && estrategicoList) {
      if (profile === "Ancla") {
        anclaList.appendChild(card);
      } else {
        estrategicoList.appendChild(card);
      }
    }
  });

  const getCardNameKey = (card) => normalizeCompanyKey(card?.querySelector(".propuesta-prospect-head strong")?.textContent || "");
  const sortCardsByOrder = (listEl, orderedKeys = []) => {
    if (!listEl || !orderedKeys.length) return;
    const indexByKey = new Map(orderedKeys.map((key, idx) => [normalizeCompanyKey(key), idx]));
    const items = Array.from(listEl.querySelectorAll(".propuesta-prospect-card"));
    items
      .sort((a, b) => {
        const aKey = getCardNameKey(a);
        const bKey = getCardNameKey(b);
        const ai = indexByKey.has(aKey) ? indexByKey.get(aKey) : Number.MAX_SAFE_INTEGER;
        const bi = indexByKey.has(bKey) ? indexByKey.get(bKey) : Number.MAX_SAFE_INTEGER;
        if (ai !== bi) return ai - bi;
        return aKey.localeCompare(bKey);
      })
      .forEach((card) => listEl.appendChild(card));
  };

  // Orden visual solicitado:
  // 1) Baja Aqua-Farms centrada sola (Ancla)
  // 2) Pinsa + GAM en la misma fila
  // 3) Pacífico + Baja Shellfish en la fila de abajo
  sortCardsByOrder(anclaList, ["Baja Aqua-Farms"]);
  sortCardsByOrder(estrategicoList, ["Grupo Pinsa", "GAM", "Pacífico Aquaculture", "Baja Shellfish Farms"]);

  // Si una lista queda con un solo prospecto (ej. Ancla), hacerla full-width y centrada.
  [anclaList, estrategicoList].forEach((listEl) => {
    if (!listEl) return;
    const totalCards = listEl.querySelectorAll(".propuesta-prospect-card").length;
    listEl.classList.toggle("propuesta-prospect-list-single", totalCards === 1);
  });
}

function syncPropuestaPlanClusters() {
  const columns = document.querySelectorAll("#tab-propuesta .propuesta-plan-column");
  if (!columns.length) return;
  const planConfig = [
    {
      badge: "Q2 2026 · mayo-junio",
      clusterTitle: "Cluster Ancla",
      empresa: "Baja Aqua-Farms",
      phases: [
        {
          title: "Entrada al mercado",
          body: "Primer cliente por criticidad térmica. Atún sashimi fresco a ≤4°C — el prospecto donde un fallo vale más que un año de servicio.",
        },
        {
          title: "Argumento de venta",
          body: "Un rechazo en Otay Mesa destruye el lote completo.",
        },
        {
          title: "Operación Baja Aqua",
          body: "88.3 cruces al mes con producto que no puede esperar. Cada viaje es una oportunidad de demostrar valor.",
        },
        {
          title: "Cierre",
          body: "Contrato piloto Q2 → contrato anual con datos reales de temperatura como evidencia.",
        },
      ],
      kpi: "88.3 DUA/mes · Otay Mesa",
    },
    {
      badge: "Q3 2026 · julio-septiembre",
      clusterTitle: "Cluster Estratégico",
      empresa: "Grupo Pinsa + GAM",
      phases: [
        {
          title: "Escalamiento comercial",
          body: "Datos reales de Baja Aqua abren la conversación. Pinsa y GAM son volumen, no urgencia.",
        },
        {
          title: "Argumento de venta",
          body: "Producto congelado = menor urgencia, pero FSMA 204 no distingue. Con 103 DUA/mes combinados, un rechazo documentado puede activar inspección sistemática en todos sus envíos.",
        },
        {
          title: "Estrategia de marketing",
          body: "ROI claro: ~$80K USD en producto perdido vs. costo anual del servicio.",
        },
        {
          title: "Cierre",
          body: "Corredor Nogales cubierto, el sustento para cerrar Q4.",
        },
      ],
      kpi: "103 DUA/mes combinados · Nogales",
    },
    {
      badge: "Q4 2026 · octubre-diciembre",
      clusterTitle: "Cluster Estratégico",
      empresa: "Pacífico + Baja Shellfish",
      phases: [
        {
          title: "Cobertura",
          body: "Los dos entran en Q4 porque necesitan ver datos reales antes de confiar un lote fresco o vivo a un sensor nuevo.",
        },
        {
          title: "Impulsor regulatorio",
          body: "Para fresco y vivo, el importador en EE.UU. exige trazabilidad térmica documentada por viaje. CL Circular lo genera automáticamente.",
        },
        {
          title: "Estrategia de marketing",
          body: "Economía circular: −50g e-waste, −200g CO₂ por chip recuperado.",
        },
        {
          title: "Cierre del año",
          body: "Con Q4 cerrado: corredor Pacífico completo, 5 contratos activos, 255.8 DUA/mes monitoreados.",
        },
      ],
      kpi: "64.7 DUA/mes combinados · Otay Mesa",
    },
  ];

  columns.forEach((column, idx) => {
    const cfg = planConfig[idx];
    if (!cfg) return;
    const badgeEl = column.querySelector(".propuesta-implementation-badge");
    const clusterTitleEl = column.querySelector("h4");
    const empresaEl = column.querySelector(".propuesta-plan-companies");
    const phaseEls = column.querySelectorAll(".propuesta-plan-phase");
    const kpiEl = column.querySelector(".propuesta-plan-kpi");

    if (badgeEl) badgeEl.textContent = cfg.badge;
    if (clusterTitleEl) clusterTitleEl.textContent = cfg.clusterTitle;
    if (empresaEl) empresaEl.textContent = cfg.empresa;
    if (kpiEl) kpiEl.textContent = cfg.kpi;

    phaseEls.forEach((phaseEl, phaseIdx) => {
      const phaseCfg = cfg.phases[phaseIdx];
      if (!phaseCfg) {
        phaseEl.hidden = true;
        return;
      }
      phaseEl.hidden = false;
      const phaseTitleEl = phaseEl.querySelector("h4");
      const phaseBodyEl = phaseEl.querySelector("p");
      if (phaseTitleEl) phaseTitleEl.textContent = phaseCfg.title;
      if (phaseBodyEl) phaseBodyEl.textContent = phaseCfg.body;
    });
  });
}

function findEmpresaByProspectName(name) {
  const target = normalizeHeader(name || "");
  if (!target) return null;

  const exact = empresasData.find((item) => normalizeHeader(item.empresa || "") === target);
  if (exact) return exact;

  if (target === "gam") {
    const gam = empresasData.find((item) => normalizeHeader(item.empresa || "").includes("grupo_acuicola_mexicano"));
    if (gam) return gam;
  }

  const loose = empresasData.find((item) => {
    const key = normalizeHeader(item.empresa || "");
    return key.includes(target) || target.includes(key);
  });
  if (loose) return loose;

  return null;
}

function findEmpresaIndexByName(name) {
  const empresa = findEmpresaByProspectName(name);
  if (!empresa) return -1;
  return empresasData.findIndex((item) => item === empresa);
}

function bindPropuestaProspectButtons() {
  const profileButtons = document.querySelectorAll("#tab-propuesta [data-prospect-company]");
  profileButtons.forEach((button) => {
    button.onclick = () => {
      const targetCompany = button.getAttribute("data-prospect-company") || "";
      openProspectProfile(targetCompany);
    };
  });

  const productButtons = document.querySelectorAll("#tab-propuesta [data-prospect-products-company]");
  productButtons.forEach((button) => {
    button.setAttribute("aria-expanded", "false");
    button.textContent = "Ver productos";
    button.onclick = () => {
      const targetCompany = button.getAttribute("data-prospect-products-company") || "";
      togglePropuestaProducts(button, targetCompany);
    };
  });
}

function togglePropuestaProducts(button, companyName) {
  const card = button?.closest(".propuesta-prospect-card");
  if (!card) return;
  const main = card.querySelector(".propuesta-prospect-main");
  if (!main) return;

  let panel = card.querySelector(".propuesta-prospect-products");
  if (!panel) {
    panel = document.createElement("div");
    panel.className = "propuesta-prospect-products";
    panel.hidden = true;
    const actions = card.querySelector(".propuesta-prospect-actions");
    if (actions && actions.parentElement === main) {
      actions.insertAdjacentElement("afterend", panel);
    } else {
      main.appendChild(panel);
    }
  }

  const isOpen = !panel.hidden;
  if (isOpen) {
    panel.hidden = true;
    button.textContent = "Ver productos";
    button.setAttribute("aria-expanded", "false");
    return;
  }

  const allPanels = document.querySelectorAll("#tab-propuesta .propuesta-prospect-products");
  allPanels.forEach((item) => {
    item.hidden = true;
  });
  const allButtons = document.querySelectorAll("#tab-propuesta [data-prospect-products-company]");
  allButtons.forEach((item) => {
    item.textContent = "Ver productos";
    item.setAttribute("aria-expanded", "false");
  });

  const empresa = findEmpresaByProspectName(companyName);
  const productos = compactRouteText(empresa?.productos || empresa?.especialidad || NO_INFO);
  panel.textContent = `Productos: ${productos}`;
  panel.hidden = false;
  button.textContent = "Ocultar productos";
  button.setAttribute("aria-expanded", "true");
}

function bindPropuestaViabilidadButton() {
  const button = document.getElementById("propuestaGoViabilidadBtn");
  if (!button) return;
  button.onclick = () => {
    const viabilidadTabBtn = document.querySelector('.tab-button[data-tab="viabilidad"]');
    if (viabilidadTabBtn) viabilidadTabBtn.click();
    setTimeout(() => {
      const shell = document.querySelector(".app-shell");
      if (shell && typeof shell.scrollIntoView === "function") {
        shell.scrollIntoView({ behavior: "smooth", block: "start" });
        return;
      }
      window.scrollTo({ top: 0, behavior: "smooth" });
    }, 0);
  };
}

function openProspectProfile(companyName) {
  const empresasTabBtn = document.querySelector('.tab-button[data-tab="empresas"]');
  if (empresasTabBtn) empresasTabBtn.click();

  const applySelection = (attempt = 0) => {
    const idx = findEmpresaIndexByName(companyName);
    if (idx < 0) {
      if (attempt < 6) setTimeout(() => applySelection(attempt + 1), 150);
      return;
    }

    if (!state.clientesShowAll) {
      state.clientesShowAll = true;
      renderEmpresas();
    }

    focusEmpresaOnMap(idx, true);
    highlightClienteRow(idx, true);
  };

  setTimeout(() => applySelection(0), 90);
}

function renderEmpresas() {
  const tbody = document.getElementById("clientesTablaBody");
  const toggleBtn = document.getElementById("clientesShowMoreBtn");
  if (!tbody) return;
  ensureClientesDuaUi();

  const visibleCount = state.clientesShowAll ? empresasData.length : Math.min(3, empresasData.length);
  const rowsToRender = empresasData.slice(0, visibleCount);

  tbody.innerHTML = rowsToRender
    .map((empresa, idx) => {
      const rutas = resolveEmpresaRutas(empresa);
      const terrestre = rutas.terrestre;
      const productos = empresa.productos || empresa.especialidad || NO_INFO;
      const certificaciones = empresa.certificaciones || NO_INFO;
      const webRaw = empresa.paginaWeb || (/^https?:\/\//i.test(empresa.contacto || "") ? empresa.contacto : "");
      const telefonoRaw =
        empresa.telefono ||
        (/(whatsapp|tel)/i.test(empresa.contacto || "") ? empresa.contacto : "");
      const emailRaw = empresa.email || extractEmailFromText(empresa.contacto || "");
      const volumenRaw = formatDuaMesLabelFromEmpresa(empresa);
      const aduanaCruceRaw = compactRouteText(empresa.cruceFronterizo || terrestre.nombre || "");
      const tempRequeridaRaw = empresa.tempRequerida || inferTempRequerida(productos, empresa.actividad || "");
      return `
      <tr class="cliente-row" data-empresa-index="${idx}">
        <td>${empresa.empresa}</td>
        <td>${productos}</td>
        <td>${formatRouteCell(certificaciones)}</td>
        <td>${formatTempCell(tempRequeridaRaw)}</td>
        <td>${formatRouteCell(volumenRaw)}</td>
        <td>${formatRouteCell(aduanaCruceRaw)}</td>
        <td>${renderContactoResumen({ telefonoRaw, emailRaw, webRaw })}</td>
      </tr>
    `;
    })
    .join("");

  initEmpresasMap();

  const rows = tbody.querySelectorAll(".cliente-row");
  rows.forEach((row) => {
    row.addEventListener("click", () => {
      const idx = Number(row.dataset.empresaIndex);
      focusEmpresaOnMap(idx, true);
    });
  });

  if (toggleBtn) {
    if (empresasData.length <= 3) {
      toggleBtn.classList.add("hidden");
    } else {
      toggleBtn.classList.remove("hidden");
      toggleBtn.textContent = state.clientesShowAll ? "Mostrar menos" : "Mostrar más";
      toggleBtn.onclick = () => {
        state.clientesShowAll = !state.clientesShowAll;
        renderEmpresas();
      };
    }
  }
}

function getEmpresaViajesAnuales(empresa) {
  const csvTrips = parseFlexibleNumber(empresa?.viajesAnuales2026 || "");
  if (Number.isFinite(csvTrips) && csvTrips > 0) return csvTrips;

  const cfgTrips = parseFlexibleNumber(getProspectMapBaseConfig(empresa?.empresa || "")?.envios || "");
  if (Number.isFinite(cfgTrips) && cfgTrips > 0) return cfgTrips;

  return NaN;
}

function formatDuaMesLabelFromEmpresa(empresa) {
  const annualTrips = getEmpresaViajesAnuales(empresa);
  if (!Number.isFinite(annualTrips) || annualTrips <= 0) return NO_INFO;
  const monthly = annualTrips / 12;
  return `${monthly.toLocaleString("es-MX", { minimumFractionDigits: 1, maximumFractionDigits: 1 })} DUA/mes`;
}

function ensureClientesDuaUi() {
  const table = document.querySelector("#tab-empresas table");
  const volumenHeader = table?.querySelector("thead th:nth-child(5)");
  if (volumenHeader) {
    volumenHeader.textContent = "Volumen (DUA/mes)";
  }

  const tableWrap = document.querySelector("#tab-empresas .clientes-table-wrap");
  if (!tableWrap) return;

  let noteEl = document.getElementById("clientesDuaNote");
  if (!noteEl) {
    noteEl = document.createElement("p");
    noteEl.id = "clientesDuaNote";
    noteEl.style.margin = "0.5rem 0 0 0";
    noteEl.style.fontSize = "0.75rem";
    noteEl.style.color = "#667784";
    tableWrap.insertAdjacentElement("afterend", noteEl);
  }
  noteEl.textContent =
    "Fuente: Veritrade — Pescados_y_Mariscos_Exportaciones.xlsx, registros aduanales 2025. DUA/mes = DUAs totales ÷ 12 meses.";
}

function initEmpresasMap() {
  const mapEl = document.getElementById("clientesMap");
  if (!mapEl || typeof L === "undefined") return;

  if (state.empresasMap) {
    state.empresasMap.remove();
    state.empresasMap = null;
    state.empresasMarkers = [];
  }

  const map = L.map(mapEl).setView([23.5, -102.5], 4.6);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);

  const bounds = [];
  empresasData.forEach((empresa) => {
    const sedeCoord = inferCoordsBySede(empresa.ubicacion || "");
    const prospectConfig = getProspectMapConfig(empresa.empresa || "");
    const lat = Number.isFinite(prospectConfig?.lat) ? prospectConfig.lat : sedeCoord.lat;
    const lng = Number.isFinite(prospectConfig?.lng) ? prospectConfig.lng : sedeCoord.lng;
    const empresaLabel = escapeHtml(compactRouteText(empresa.empresa || "Empresa"));
    const perfil = escapeHtml(prospectConfig?.perfil || "Estratégico");
    const envios = escapeHtml(
      formatEnviosLabel(empresa.viajesAnuales2026 || prospectConfig?.envios || empresa.volumenEstimado || ""),
    );
    const producto = escapeHtml(compactRouteText(empresa.productos || empresa.especialidad || NO_INFO));
    const trimestre = escapeHtml(prospectConfig?.trimestre || "Q3 2026");
    const popup = `
      <div class="prospect-popup">
        <strong>${empresaLabel}</strong><br/>
        <span><strong>Perfil:</strong> ${perfil}</span><br/>
        <span><strong>Envíos/año:</strong> ${envios}</span><br/>
        <span><strong>Producto principal:</strong> ${producto}</span><br/>
        <span><strong>Trimestre de contacto:</strong> ${trimestre}</span>
      </div>
    `;

    const marker = L.marker([lat, lng], {
      icon: buildProspectProfileIcon(prospectConfig?.perfil || "Estratégico"),
    })
      .addTo(map)
      .bindPopup(popup);

    state.empresasMarkers.push(marker);
    bounds.push([lat, lng]);
  });

  if (bounds.length) {
    map.fitBounds(bounds, { padding: [90, 90], maxZoom: 5 });
  }

  addProspectMapLegend(map);

  state.empresasMap = map;
}

function getProspectMapBaseConfig(name = "") {
  const key = normalizeCompanyKey(name);
  if (key.includes("GRUPO PINSA")) {
    return {
      lat: 23.2494,
      lng: -106.4111,
      markerCode: "A1",
      envios: "513",
      trimestre: "Q2 2026",
    };
  }
  if (key.includes("GRUPO ACUICOLA MEXICANO") || key === "GAM" || key.includes(" GAM")) {
    return {
      lat: 23.2194,
      lng: -106.4411,
      markerCode: "A2",
      envios: "720",
      trimestre: "Q2 2026",
    };
  }
  if (key.includes("BAJA AQUA FARMS")) {
    return {
      lat: 31.88,
      lng: -116.59,
      markerCode: "E1",
      envios: "1060",
      trimestre: "Q3 2026",
    };
  }
  if (key.includes("PACIFICO AQUACULTURE")) {
    return {
      lat: 31.84,
      lng: -116.62,
      markerCode: "E2",
      envios: "296",
      trimestre: "Q3 2026",
    };
  }
  if (key.includes("BAJA SHELLFISH FARMS")) {
    return {
      lat: 31.86,
      lng: -116.65,
      markerCode: "E3",
      envios: "480",
      trimestre: "Q4 2026",
    };
  }
  return null;
}

function getProspectMapConfig(name = "") {
  const base = getProspectMapBaseConfig(name);
  if (!base) return null;
  const clusterRow = getClusteredProspectRowByName(name);
  return {
    ...base,
    perfil: clusterRow?.profile || "Estratégico",
    trimestre: clusterRow?.quarter || base.trimestre || "Q3 2026",
  };
}

function buildProspectMarkerIcon(code, perfil = "Estratégico") {
  const profileClass = perfil === "Ancla" ? "prospect-marker-ancla" : "prospect-marker-estrategico";
  return L.divIcon({
    className: "prospect-marker-wrap",
    html: `<div class="prospect-marker ${profileClass}">${escapeHtml(code)}</div>`,
    iconSize: [34, 34],
    iconAnchor: [17, 17],
    popupAnchor: [0, -16],
  });
}

function addProspectMapLegend(map) {
  const legend = L.control({ position: "bottomleft" });
  legend.onAdd = () => {
    const div = L.DomUtil.create("div", "prospect-map-legend");
    div.innerHTML = `
      <div class="prospect-map-legend-item">
        <span class="prospect-map-legend-dot prospect-map-legend-dot-ancla"></span>
        Verde oscuro = Cluster Ancla
      </div>
      <div class="prospect-map-legend-item">
        <span class="prospect-map-legend-dot prospect-map-legend-dot-estrategico"></span>
        Verde medio = Cluster Estratégico
      </div>
    `;
    return div;
  };
  legend.addTo(map);
}

function formatFtlLabelFromVolume(value = "") {
  const text = compactRouteText(value);
  if (!text || text === NO_INFO) return NO_INFO;
  const match = text.match(/(\d+\s*(?:-\s*\d+)?)\s*FTL/i);
  if (match && match[1]) return match[1].replace(/\s+/g, "");
  return text;
}

function formatEnviosLabel(value = "") {
  const text = compactRouteText(value);
  if (!text || text === NO_INFO) return NO_INFO;
  if (/env[ií]os|viajes/i.test(text)) return text;

  const rangeMatch = text.match(/(\d+(?:[.,]\d+)?)\s*(?:a|-|–|—)\s*(\d+(?:[.,]\d+)?)/i);
  if (rangeMatch) {
    const min = Number(rangeMatch[1].replace(/,/g, ""));
    const max = Number(rangeMatch[2].replace(/,/g, ""));
    if (Number.isFinite(min) && Number.isFinite(max)) {
      return `${Math.round(min).toLocaleString("es-MX")}-${Math.round(max).toLocaleString("es-MX")} envíos/año`;
    }
  }

  const numberMatch = text.match(/(\d+(?:[.,]\d+)?)/);
  if (numberMatch) {
    const valueNum = Number(numberMatch[1].replace(/,/g, ""));
    if (Number.isFinite(valueNum)) return `${Math.round(valueNum).toLocaleString("es-MX")} envíos/año`;
  }

  return `${text} envíos/año`;
}

function focusEmpresaOnMap(index, openPopup = false) {
  const marker = state.empresasMarkers[index];
  if (!marker || !state.empresasMap) return;
  const currentZoom = state.empresasMap.getZoom();
  const targetZoom = Math.min(Math.max(currentZoom, 5), 6);
  state.empresasMap.flyTo(marker.getLatLng(), targetZoom, { duration: 0.5 });
  if (openPopup) marker.openPopup();
  highlightClienteRow(index, false);
}

function highlightClienteRow(index, shouldScroll = false) {
  const rows = document.querySelectorAll(".cliente-row");
  let selected = null;
  rows.forEach((row, idx) => {
    const rowIndex = Number(row.dataset.empresaIndex);
    const isSelected = Number.isFinite(rowIndex) ? rowIndex === index : idx === index;
    row.classList.toggle("is-selected", isSelected);
    if (isSelected) selected = row;
  });
  if (shouldScroll && selected) {
    selected.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
  }
}

function renderContacto(empresa) {
  if (!empresa.contactoLink) return empresa.contacto;
  return `<a href="${empresa.contactoLink}" target="_blank" rel="noopener noreferrer">${empresa.contacto}</a>`;
}

function renderWebsite(raw) {
  const links = splitLinks(raw);
  if (!links.length) return NO_INFO;
  return links
    .map((link) => `<a href="${link}" target="_blank" rel="noopener noreferrer">${shortLabel(link)}</a>`)
    .join("<br/>");
}

function renderTelefono(raw) {
  const text = String(raw || "").trim();
  if (!text) return NO_INFO;
  const phoneDigits = extractFirstPhoneDigits(text);
  if (phoneDigits) {
    return `<a href="https://wa.me/${phoneDigits}" target="_blank" rel="noopener noreferrer">${text}</a>`;
  }
  return text;
}

function renderEmail(raw) {
  const text = String(raw || "").trim();
  if (!text) return NO_INFO;
  const first = text.split(",")[0].trim().replace(/\.$/, "");
  return `<a href="mailto:${first}">${first}</a>`;
}

function renderContactoResumen({ telefonoRaw = "", emailRaw = "", webRaw = "" } = {}) {
  const tel = String(telefonoRaw || "").trim();
  const emailMain = String(emailRaw || "").trim();
  const emailFromWeb = extractEmailFromText(String(webRaw || ""));
  const email = emailMain || emailFromWeb;
  const linksWeb = splitLinks(webRaw);

  const lines = [];
  if (tel) lines.push(`<span><strong>Tel:</strong> ${renderTelefono(tel)}</span>`);
  if (email) lines.push(`<span><strong>Email:</strong> ${renderEmail(email)}</span>`);
  if (linksWeb.length) lines.push(`<span><strong>Web:</strong> ${renderWebsite(webRaw)}</span>`);

  if (!lines.length) return NO_INFO;
  return lines.join("<br/>");
}

function shortLabel(link) {
  try {
    const url = new URL(link);
    return url.hostname.replace(/^www\./, "");
  } catch (error) {
    return "Sitio";
  }
}

function extractEmailFromText(text) {
  const match = String(text || "").match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return match ? match[0] : "";
}

function toMapsLink(origen, destino) {
  const query = encodeURIComponent(`${origen} a ${destino}`);
  return `https://www.google.com/maps/search/?api=1&query=${query}`;
}

function getRutaTerrestre(ubicacion) {
  const u = ubicacion.toLowerCase();
  if (u.includes("baja california") || u.includes("ensenada") || u.includes("tijuana")) {
    return {
      nombre: "Tijuana (Baja California) - San Ysidro (California)",
      razon: "Cruce más directo del corredor oeste para carga refrigerada.",
    };
  }
  if (u.includes("yucatan") || u.includes("quintana roo") || u.includes("veracruz")) {
    return {
      nombre: "Reynosa (Tamaulipas) - Pharr/McAllen (Texas)",
      razon: "Hub logístico con alto flujo de perecederos del sureste y golfo.",
    };
  }
  if (u.includes("chihuahua")) {
    return {
      nombre: "Ciudad Juárez (Chihuahua) - El Paso (Texas)",
      razon: "Corredor estratégico para el centro-norte y frontera norte.",
    };
  }
  return {
    nombre: "Nuevo Laredo (Tamaulipas) - Laredo (Texas)",
    razon: "Principal corredor terrestre México - Estados Unidos por volumen comercial.",
  };
}

function getRutaOceanica(ubicacion) {
  const u = ubicacion.toLowerCase();
  if (u.includes("baja california") || u.includes("ensenada")) {
    return {
      nombre: "Ensenada (Baja California)",
      razon: "Puerto especializado en productos pesqueros del noroeste.",
    };
  }
  if (u.includes("sinaloa") || u.includes("sonora")) {
    return {
      nombre: "Mazatlán (Sinaloa)",
      razon: "Puerto pesquero y comercial con fuerte tradición camaronera.",
    };
  }
  if (u.includes("jalisco") || u.includes("colima")) {
    return {
      nombre: "Manzanillo (Colima)",
      razon: "Mayor movimiento de contenedores del Pacífico mexicano.",
    };
  }
  if (u.includes("yucatan")) {
    return {
      nombre: "Progreso (Yucatán)",
      razon: "Puerto estratégico del sureste para salida de perecederos.",
    };
  }
  if (u.includes("quintana roo") || u.includes("veracruz") || u.includes("tamaulipas")) {
    return {
      nombre: "Veracruz (Veracruz)",
      razon: "Principal puerta del Golfo para comercio marítimo.",
    };
  }
  return {
    nombre: "Lázaro Cárdenas (Michoacán)",
    razon: "Puerto de aguas profundas con alta capacidad para volúmenes grandes.",
  };
}

function getEmpresaRouteCatalog(empresaNombre) {
  const target = normalizeGeoKey(empresaNombre);
  return Object.entries(empresaRouteCatalog).find(([name]) => normalizeGeoKey(name) === target)?.[1] || null;
}

function getEmpresaRutaTerrestreRaw(empresa, fallback = "") {
  const catalog = getEmpresaRouteCatalog(empresa?.empresa || "");
  return String(empresa?.rutaTerrestre || empresa?.cruceFronterizo || catalog?.rutaTerrestre || fallback || "").trim();
}

function getEmpresaRutaMaritimaRaw(empresa, fallback = "") {
  const catalog = getEmpresaRouteCatalog(empresa?.empresa || "");
  return String(empresa?.rutaMaritima || catalog?.rutaMaritima || fallback || "").trim();
}

function resolveEmpresaRutas(empresa) {
  const fallbackTerrestre = getRutaTerrestre(empresa.ubicacion || "");
  const fallbackOceanica = getRutaOceanica(empresa.ubicacion || "");
  const terrestreRaw = getEmpresaRutaTerrestreRaw(empresa, fallbackTerrestre.nombre);
  const oceanicaRaw = getEmpresaRutaMaritimaRaw(empresa, fallbackOceanica.nombre);
  const terrestreText = normalizeGeoKey(`${empresa.cruceFronterizo || ""} ${terrestreRaw}`);
  const oceanicaText = normalizeGeoKey(oceanicaRaw);

  const terrestre = detectTerrestreByText(terrestreText) || fallbackTerrestre;
  const oceanica = detectOceanicaByText(oceanicaText) || fallbackOceanica;
  return { terrestre, oceanica };
}

function detectTerrestreByText(text) {
  if (!text) return null;
  const options = [
    {
      nombre: "Nuevo Laredo (Tamaulipas) - Laredo (Texas)",
      keys: ["NUEVO LAREDO", "LAREDO"],
    },
    {
      nombre: "Reynosa (Tamaulipas) - Pharr/McAllen (Texas)",
      keys: ["REYNOSA", "PHARR", "MCALLEN"],
    },
    {
      nombre: "Tijuana (Baja California) - San Ysidro (California)",
      keys: ["TIJUANA", "SAN YSIDRO", "OTAY", "TECATE"],
    },
    {
      nombre: "Ciudad Juárez (Chihuahua) - El Paso (Texas)",
      keys: ["JUAREZ", "EL PASO"],
    },
    {
      nombre: "Matamoros (Tamaulipas) - Brownsville (Texas)",
      keys: ["MATAMOROS", "BROWNSVILLE"],
    },
    {
      nombre: "Nogales (Sonora) - Nogales (Arizona)",
      keys: ["NOGALES"],
    },
  ];
  let best = null;
  let bestIndex = Number.POSITIVE_INFINITY;
  options.forEach((opt) => {
    opt.keys.forEach((k) => {
      const idx = text.indexOf(k);
      if (idx >= 0 && idx < bestIndex) {
        best = { nombre: opt.nombre };
        bestIndex = idx;
      }
    });
  });
  return best;
}

function detectOceanicaByText(text) {
  if (!text) return null;
  const options = [
    { nombre: "Ensenada (Baja California)", key: "ENSENADA" },
    { nombre: "Mazatlán (Sinaloa)", key: "MAZATLAN" },
    { nombre: "Manzanillo (Colima)", key: "MANZANILLO" },
    { nombre: "Lázaro Cárdenas (Michoacán)", key: "LAZARO CARDENAS" },
    { nombre: "Veracruz (Veracruz)", key: "VERACRUZ" },
    { nombre: "Altamira (Tamaulipas)", key: "ALTAMIRA" },
    { nombre: "Tampico (Tamaulipas)", key: "TAMPICO" },
    { nombre: "Progreso (Yucatán)", key: "PROGRESO" },
  ];
  let best = null;
  let bestIndex = Number.POSITIVE_INFINITY;
  options.forEach((opt) => {
    const idx = text.indexOf(opt.key);
    if (idx >= 0 && idx < bestIndex) {
      best = { nombre: opt.nombre };
      bestIndex = idx;
    }
  });
  return best;
}

function hasMeaningfulRoute(text) {
  const value = normalizeGeoKey(text);
  if (!value) return false;
  if (["N D", "NA", "NO DISPONIBLE", "NO DISPONIBLE PUBLICAMENTE"].includes(value)) return false;
  if (value.includes("NO APLICA")) return false;
  if (value.includes("SIN EXPORTACION MARITIMA")) return false;
  if (value.includes("SIN CRUCE FRONTERIZO ACTIVO")) return false;
  if (value.includes("DISTRIBUCION NACIONAL") && value.includes("SIN")) return false;
  return true;
}

function findBestAliasInText(text, aliases) {
  const normalized = normalizeGeoKey(text);
  if (!normalized) return null;
  let best = null;
  let bestLength = -1;
  let bestIndex = Number.POSITIVE_INFINITY;

  aliases.forEach((alias) => {
    alias.keys.forEach((key) => {
      const idx = normalized.indexOf(key);
      if (idx < 0) return;
      if (key.length > bestLength || (key.length === bestLength && idx < bestIndex)) {
        best = alias;
        bestLength = key.length;
        bestIndex = idx;
      }
    });
  });

  if (!best) return null;
  return { label: best.label, lat: best.lat, lng: best.lng };
}

function extractWaypointsByRouteOrder(rawRoute, aliases) {
  const routeRaw = String(rawRoute || "").trim();
  if (!routeRaw) return [];
  const segments = routeRaw
    .split(/\s*(?:->|→)\s*/)
    .map((segment) => segment.trim())
    .filter(Boolean);
  if (segments.length < 2) return [];

  const points = [];
  segments.forEach((segment) => {
    const alternatives = segment
      .split(/\s*\/\s*/)
      .map((part) => part.trim())
      .filter(Boolean);

    let match = null;
    for (const alt of alternatives) {
      match = findBestAliasInText(alt, aliases);
      if (match) break;
    }
    if (!match) {
      match = findBestAliasInText(segment, aliases);
    }
    if (!match) return;

    const prev = points[points.length - 1];
    if (prev && prev.label === match.label) return;
    points.push(match);
  });

  return points.length >= 2 ? points : [];
}

function extractGeoWaypointsFromRoute(rawRoute, mode = "terrestre") {
  if (!hasMeaningfulRoute(rawRoute)) return [];
  const aliases = routeWaypointAliases.filter((item) => item.modes.includes(mode));
  const byOrder = extractWaypointsByRouteOrder(rawRoute, aliases);
  if (byOrder.length >= 2) return byOrder;

  const text = normalizeGeoKey(rawRoute);
  const matches = [];

  aliases.forEach((alias) => {
    alias.keys.forEach((key) => {
      let idx = text.indexOf(key);
      while (idx >= 0) {
        matches.push({
          idx,
          keyLength: key.length,
          label: alias.label,
          lat: alias.lat,
          lng: alias.lng,
        });
        idx = text.indexOf(key, idx + key.length);
      }
    });
  });

  matches.sort((a, b) => (a.idx === b.idx ? b.keyLength - a.keyLength : a.idx - b.idx));
  const deduped = [];
  matches.forEach((match) => {
    const prev = deduped[deduped.length - 1];
    if (prev && prev.label === match.label) return;
    if (prev && prev._idx === match.idx && prev._keyLength > match.keyLength) return;
    deduped.push({
      label: match.label,
      lat: match.lat,
      lng: match.lng,
      _idx: match.idx,
      _keyLength: match.keyLength,
    });
  });

  return deduped.map(({ label, lat, lng }) => ({ label, lat, lng }));
}

function compactRouteText(text) {
  return String(text || "")
    .replace(/\s*\n+\s*/g, " | ")
    .replace(/\s+/g, " ")
    .trim();
}

function toHtmlMultiline(text) {
  return escapeHtml(String(text || "").trim()).replace(/\n/g, "<br/>");
}

function formatRouteCell(text) {
  const value = String(text || "").trim();
  if (!value) return NO_INFO;
  return escapeHtml(value).replace(/\n/g, "<br/>");
}

function formatTempCell(text) {
  const value = formatTempDisplayText(text, { includeFrozenAlso: true });
  if (!value) return NO_INFO;
  const firstLine = value.split("\n").map((part) => part.trim()).find(Boolean) || value;
  const full = compactRouteText(value);
  return `<span title="${escapeHtml(full)}">${escapeHtml(firstLine)}</span>`;
}

function formatTempDisplayText(text, { includeFrozenAlso = false } = {}) {
  let value = compactRouteText(stripNom242Tag(text || ""));
  if (!value) return "";

  value = value
    .replace(/[−–—]/g, "-")
    .replace(/≤/g, "<=")
    .replace(/≥/g, ">=")
    .replace(/\bCong(?:\.|elado)?\s*:/gi, "Congelado:")
    .replace(/\bTemp\.\s*ambiente\b/gi, "Temp. ambiente")
    .replace(/°\s*C/gi, " C")
    .replace(/\s*\/\s*/g, " | ")
    .replace(/\s*·\s*/g, " | ")
    .replace(/<=\s*/g, "<= ")
    .replace(/>=\s*/g, ">= ")
    .replace(/\s{2,}/g, " ")
    .trim();

  const normalized = normalizeTempRiskText(value);
  const hasFrozenMarker =
    normalized.includes("congelado") || hasTempMarker(normalized, "-18") || hasTempMarker(normalized, "-60");

  if (includeFrozenAlso && !hasFrozenMarker) {
    value = `${value} | Congelado: <= -18 C también`;
  }

  return value;
}

function escapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

async function loadWorkbookFromPath(path) {
  const response = await fetch(path);
  if (!response.ok) {
    throw new Error(`No se pudo leer ${path}`);
  }
  const arrayBuffer = await response.arrayBuffer();
  return XLSX.read(arrayBuffer, { type: "array" });
}

function initFileUpload() {
  const input = document.getElementById("xlsxInput");
  if (!input) return;
  input.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    let attempts = 0;
    while (typeof XLSX === "undefined" && attempts < 20) {
      await new Promise((resolve) => setTimeout(resolve, 200));
      attempts += 1;
    }

    try {
      ensureLibs();
      const workbook = await loadWorkbookFromFile(file);
      loadDataAndRender(workbook);
      const mercadoTabBtn = document.querySelector('.tab-button[data-tab="kpis"]');
      if (mercadoTabBtn && !mercadoTabBtn.classList.contains("is-active")) mercadoTabBtn.click();
      setStatus(`Archivo cargado: ${file.name}`, true);
      alert(`Archivo cargado correctamente: ${file.name}`);
    } catch (error) {
      setStatus("No se pudo leer el archivo seleccionado.", false);
      alert(`No se pudo procesar el XLSX: ${error?.message || "revisa hojas y columnas."}`);
      console.error(error);
    } finally {
      input.value = "";
    }
  });
}

function initCsvUpload() {
  const input = document.getElementById("csvInput");
  if (!input) return;
  input.addEventListener("change", async (event) => {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;

    const changed = {
      empresas: false,
      competidores: false,
      exportaciones: false,
      cruces: false,
      puertos: false,
    };
    const loaded = [];
    const failed = [];

    try {
      for (const file of files) {
        try {
          const text = await file.text();
          const target = detectCsvTarget(file.name, text);
          if (!target) {
            failed.push(`${file.name} (tipo no reconocido)`);
            continue;
          }

          if (target === "empresas") {
            const parsed = parseEmpresasCsv(text);
            if (!parsed.length) throw new Error("sin filas válidas");
            empresasData = parsed;
            state.empresasSource = "manual";
            state.empresasHash = buildEmpresasHash(empresasData);
            changed.empresas = true;
            loaded.push(`${file.name} -> Empresas`);
            continue;
          }

          if (target === "competidores") {
            const parsed = parseCompetidoresCsv(text);
            if (!parsed.length) throw new Error("sin filas válidas");
            competidoresData = parsed;
            changed.competidores = true;
            loaded.push(`${file.name} -> Competidores`);
            continue;
          }

          if (target === "exportaciones") {
            const parsed = parseExportacionesFobCsv(text);
            if (!parsed.length) throw new Error("sin filas válidas");
            state.exportFobSeries = parsed;
            state.exportFobSource = "manual";
            changed.exportaciones = true;
            loaded.push(`${file.name} -> Exportaciones FOB`);
            continue;
          }

          if (target === "cruces") {
            const parsed = parseCrucesTerrestresCsv(text);
            if (!parsed.length) throw new Error("sin filas válidas");
            state.infraCruces = parsed;
            state.infraSource = "manual";
            changed.cruces = true;
            loaded.push(`${file.name} -> Cruces terrestres`);
            continue;
          }

          if (target === "puertos") {
            const parsed = parsePuertosOceanicosCsv(text);
            if (!parsed.length) throw new Error("sin filas válidas");
            state.infraPuertos = parsed;
            state.infraSource = "manual";
            changed.puertos = true;
            loaded.push(`${file.name} -> Puertos oceánicos`);
          }
        } catch (error) {
          failed.push(`${file.name} (${error?.message || "error de parseo"})`);
        }
      }

      if (changed.empresas) {
        renderEmpresas();
        renderClustering();
        syncRiesgoEmpresaOptions();
      }
      if (changed.cruces || changed.puertos) {
        initInfraKpi();
        const select = document.getElementById("riesgoEmpresaSelect");
        if (select && empresasData.length) {
          updateRiesgosByEmpresa(Number(select.value || 0));
        }
      }
      if (changed.empresas || changed.exportaciones) {
        renderAll();
      } else {
        if (changed.competidores) renderCompetidoresKpi();
        if (changed.exportaciones) renderSerieCharts();
      }
      if (changed.empresas) {
        const select = document.getElementById("riesgoEmpresaSelect");
        if (select && empresasData.length) {
          updateRiesgosByEmpresa(Number(select.value || 0));
        }
      }

      const messageParts = [];
      if (loaded.length) messageParts.push(`CSV cargados:\n- ${loaded.join("\n- ")}`);
      if (failed.length) messageParts.push(`No cargados:\n- ${failed.join("\n- ")}`);
      if (messageParts.length) alert(messageParts.join("\n\n"));
    } finally {
      input.value = "";
    }
  });
}

function detectCsvTarget(fileName = "", csvText = "") {
  const normalizedName = normalizeHeader(fileName || "");
  if (normalizedName.includes("empresas")) return "empresas";
  if (normalizedName.includes("competidores")) return "competidores";
  if (normalizedName.includes("exportaciones") || normalizedName.includes("fob")) return "exportaciones";
  if (normalizedName.includes("cruces")) return "cruces";
  if (normalizedName.includes("puertos")) return "puertos";

  const rows = parseCsvText(csvText);
  const headerRow = rows.find((row) => Array.isArray(row) && row.some((cell) => String(cell || "").trim())) || [];
  const headers = headerRow.map((h) => normalizeHeader(h));
  const hasHeader = (needle) => {
    const norm = normalizeHeader(needle);
    return headers.some((h) => h === norm || h.startsWith(norm));
  };

  if (hasHeader("empresa") && (hasHeader("temp_requerida") || hasHeader("ruta_terrestre") || hasHeader("certificaciones"))) {
    return "empresas";
  }
  if (hasHeader("empresa") && (hasHeader("servicio_principal") || hasHeader("sede_en_mexico") || hasHeader("propuesta_de_valor"))) {
    return "competidores";
  }
  if ((hasHeader("ano") || hasHeader("anio") || hasHeader("year")) && (hasHeader("exportaciones_fob") || hasHeader("fob") || hasHeader("valor"))) {
    return "exportaciones";
  }
  if (hasHeader("aduana_mx") || (hasHeader("tiempo_cruce_con_inspeccion_fda_hrs") && hasHeader("ftl_mariscos_mes_est"))) {
    return "cruces";
  }
  if (hasHeader("puerto") && (hasHeader("latitud") || hasHeader("longitud"))) {
    return "puertos";
  }
  return "";
}

async function loadWorkbookFromFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  return XLSX.read(arrayBuffer, { type: "array" });
}

function loadDataAndRender(workbook) {
  const backup = {
    rows2024: Array.isArray(state.rows2024) ? [...state.rows2024] : [],
    rows2023: Array.isArray(state.rows2023) ? [...state.rows2023] : [],
    rowsByYear: Object.fromEntries(
      Object.entries(state.rowsByYear || {}).map(([year, rows]) => [year, Array.isArray(rows) ? [...rows] : []]),
    ),
    yearsAvailable: Array.isArray(state.yearsAvailable) ? [...state.yearsAvailable] : [],
    resumen: Array.isArray(state.resumen) ? [...state.resumen] : [],
    entidades2024: Array.isArray(state.entidades2024) ? [...state.entidades2024] : [],
    speciesCaptura: Array.isArray(state.speciesCaptura) ? [...state.speciesCaptura] : [],
    speciesAcuacultura: Array.isArray(state.speciesAcuacultura) ? [...state.speciesAcuacultura] : [],
  };

  state.rows2024 = [];
  state.rows2023 = [];
  state.rowsByYear = {};
  state.yearsAvailable = [];
  state.resumen = [];
  state.entidades2024 = [];
  state.speciesCaptura = [];
  state.speciesAcuacultura = [];
  try {
    processWorkbook(workbook);
    renderAll();
  } catch (error) {
    state.rows2024 = backup.rows2024;
    state.rows2023 = backup.rows2023;
    state.rowsByYear = backup.rowsByYear;
    state.yearsAvailable = backup.yearsAvailable;
    state.resumen = backup.resumen;
    state.entidades2024 = backup.entidades2024;
    state.speciesCaptura = backup.speciesCaptura;
    state.speciesAcuacultura = backup.speciesAcuacultura;
    renderAll();
    throw error;
  }
}

function processWorkbook(workbook) {
  const sheetRefs = resolveWorkbookSheets(workbook);
  const baseSheet = sheetRefs.baseSheet;
  const resumenSheet = sheetRefs.resumenSheet;

  if (!baseSheet && !resumenSheet) {
    throw new Error("El libro no contiene hojas válidas (BASE_COMPLETA o RESUMEN_ANUAL).");
  }

  const baseRows = baseSheet
    ? XLSX.utils.sheet_to_json(baseSheet, { header: 1, raw: true, defval: null, blankrows: false })
    : [];
  const resumenRows = resumenSheet
    ? XLSX.utils.sheet_to_json(resumenSheet, { header: 1, raw: true, defval: null, blankrows: false })
    : [];

  state.resumen = resumenRows;
  console.log("CONAPESCA base rows:", baseRows.length);
  console.log("CONAPESCA resumen rows:", resumenRows.length);
  if (baseRows.length) console.log("Sample BASE row:", baseRows[0]);

  const speciesCapturaSet = new Set();
  const speciesAcuaculturaSet = new Set();
  const entidadesSet = new Set();
  const yearsSet = new Set();

  const addParsedRow = (year, parsed) => {
    if (!state.rowsByYear[year]) state.rowsByYear[year] = [];
    state.rowsByYear[year].push(parsed);
    yearsSet.add(year);

    if (year === 2024) {
      state.rows2024.push(parsed);
      entidadesSet.add(parsed.entidad);
      if (parsed.origen === "CAPTURA") speciesCapturaSet.add(parsed.especie);
      if (parsed.origen === "ACUACULTURA") speciesAcuaculturaSet.add(parsed.especie);
    } else if (year === 2023) {
      state.rows2023.push(parsed);
    }
  };

  const requiredCols = [
    "ANO",
    "ORIGEN",
    "ESPECIE",
    "ENTIDAD",
    "MES",
    "LITORAL",
    "PESO_DESEMBARCADO_KG",
    "VALOR_PESOS",
  ];
  if (baseRows.length) {
    const headerRowIndex = findHeaderRowIndex(baseRows, requiredCols);
    if (headerRowIndex < 0) {
      const fallbackRows = parseBaseCompletaFallbackRows(baseRows);
      fallbackRows.forEach((parsed) => {
        if (!Number.isFinite(parsed.year)) return;
        addParsedRow(parsed.year, parsed);
      });
    } else {
      const headers = baseRows[headerRowIndex] || [];
      const idx = indexMap(headers);
      const missing = requiredCols.filter((key) => !Number.isInteger(idx[key]));
      if (missing.length) {
        const fallbackRows = parseBaseCompletaFallbackRows(baseRows);
        fallbackRows.forEach((parsed) => {
          if (!Number.isFinite(parsed.year)) return;
          addParsedRow(parsed.year, parsed);
        });
      } else {
        for (let i = headerRowIndex + 1; i < baseRows.length; i += 1) {
          const r = baseRows[i];
          if (!r || !r.some((v) => String(v ?? "").trim())) continue;
          const year = Number(r?.[idx.ANO]);
          if (!r || !Number.isFinite(year)) continue;

          const origen = String(r[idx.ORIGEN] || "").trim().toUpperCase();
          const especie = String(r[idx.ESPECIE] || "SIN ESPECIE").trim().toUpperCase();
          const entidad = String(r[idx.ENTIDAD] || "SIN ENTIDAD").trim().toUpperCase();
          const mes = String(r[idx.MES] || "SIN MES").trim().toUpperCase();
          const litoral = String(r[idx.LITORAL] || "SIN DATO").trim().toUpperCase();
          const pesoKg = Number(r[idx.PESO_DESEMBARCADO_KG]) || 0;
          const valorPesos = Number(r[idx.VALOR_PESOS]) || 0;

          const parsed = {
            origen,
            especie,
            entidad,
            mes,
            litoral,
            pesoKg,
            valorPesos,
          };
          addParsedRow(year, parsed);
        }
      }
    }
  }

  // Fallback adicional: si BASE no trae filas útiles, usa RESUMEN_ANUAL por columnas.
  if (!yearsSet.size) {
    const resumenFallbackRows = parseResumenAnualFallbackRows(resumenRows);
    resumenFallbackRows.forEach((parsed) => {
      if (!Number.isFinite(parsed.year)) return;
      addParsedRow(parsed.year, parsed);
    });
  }

  state.yearsAvailable = Array.from(yearsSet).sort((a, b) => b - a);
  state.entidades2024 = Array.from(entidadesSet).sort();
  state.speciesCaptura = Array.from(speciesCapturaSet).sort();
  state.speciesAcuacultura = Array.from(speciesAcuaculturaSet).sort();
  if (!state.yearsAvailable.length) {
    throw new Error("No se encontraron datos. Verifica que el archivo es CONAPESCA BASE_COMPLETA.");
  }
}

function parseBaseCompletaFallbackRows(baseRows = []) {
  const parsed = [];
  for (let i = 0; i < baseRows.length; i += 1) {
    const r = baseRows[i] || [];
    if (!r.some((v) => String(v ?? "").trim())) continue;
    const year = Number(r[0]);
    if (!Number.isFinite(year)) continue;
    const pesoKg = Number(r[11]) || 0;
    const valorPesos = Number(r[13]) || 0;
    if (pesoKg === 0 && valorPesos === 0) continue;
    parsed.push({
      year,
      origen: "CAPTURA",
      especie: "TOTAL",
      entidad: "NACIONAL",
      mes: "ANUAL",
      litoral: "SIN DATO",
      pesoKg,
      valorPesos,
    });
  }
  return parsed;
}

function parseResumenAnualFallbackRows(resumenRows = []) {
  const parsed = [];
  for (let i = 0; i < resumenRows.length; i += 1) {
    const r = resumenRows[i] || [];
    if (!r.some((v) => String(v ?? "").trim())) continue;
    const year = Number(r[0]);
    if (!Number.isFinite(year)) continue;

    // Estructura observada en RESUMEN_ANUAL sin encabezados:
    // A=ANO, C=PESO_DESEMBARCADO_TON, E=VALOR_MILLONES_MXN.
    const pesoTon = Number(r[2]);
    const valorMillonesMxn = Number(r[4]);
    if (!Number.isFinite(pesoTon) && !Number.isFinite(valorMillonesMxn)) continue;

    parsed.push({
      year,
      origen: "TOTAL",
      especie: "TOTAL",
      entidad: "NACIONAL",
      mes: "ANUAL",
      litoral: "SIN DATO",
      pesoKg: (Number.isFinite(pesoTon) ? pesoTon : 0) * 1000,
      valorPesos: (Number.isFinite(valorMillonesMxn) ? valorMillonesMxn : 0) * 1_000_000,
    });
  }
  return parsed;
}

function findHeaderRowIndex(rows = [], requiredCols = []) {
  if (!Array.isArray(rows) || !rows.length) return -1;
  const required = requiredCols.map((col) => normalizeHeader(col).toUpperCase());

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i] || [];
    if (!row.some((cell) => String(cell ?? "").trim())) continue;
    const map = indexMap(row);
    const hasAll = required.every((key) => Number.isInteger(map[key]));
    if (hasAll) return i;
  }
  return -1;
}

function resolveWorkbookSheets(workbook) {
  const sheetNames = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!sheetNames.length) {
    return { baseSheet: null, resumenSheet: null };
  }

  const normalizedByName = new Map();
  sheetNames.forEach((name) => {
    normalizedByName.set(name, normalizeHeader(name));
  });

  const findSheetName = (predicates = []) =>
    sheetNames.find((name) => {
      const normalized = normalizedByName.get(name) || "";
      return predicates.some((predicate) => predicate(normalized));
    });

  const baseName =
    findSheetName([
      (n) => n === "base_completa",
      (n) => n.startsWith("base_completa_"),
      (n) => n.includes("base_completa"),
      (n) => n === "base",
      (n) => n.startsWith("base_"),
    ]) || "BASE_COMPLETA";

  const resumenName =
    findSheetName([
      (n) => n === "resumen_anual",
      (n) => n.startsWith("resumen_anual_"),
      (n) => n.includes("resumen_anual"),
      (n) => n === "resumen",
      (n) => n.startsWith("resumen_"),
    ]) || "RESUMEN_ANUAL";

  return {
    baseSheet: workbook.Sheets?.[baseName] || null,
    resumenSheet: workbook.Sheets?.[resumenName] || null,
  };
}

function indexMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const raw = String(h || "").trim();
    if (!raw) return;
    map[raw] = i;
    map[normalizeHeader(raw).toUpperCase()] = i;
  });
  return map;
}

function renderAll() {
  if (!Array.isArray(state.yearsAvailable) || !state.yearsAvailable.length || !Object.keys(state.rowsByYear || {}).length) {
    seedDefaultKpiData();
  }
  try {
    initFilters();
  } catch (error) {
    console.error("initFilters error:", error);
  }
  try {
    renderKpiCards();
  } catch (error) {
    console.error("renderKpiCards error:", error);
    renderKpiCardsEmergencyFallback();
  }
  try {
    renderMapaEstados();
  } catch (error) {
    console.error("renderMapaEstados error:", error);
  }
  if (!Array.isArray(state.exportFobSeries) || !state.exportFobSeries.length) {
    state.exportFobSeries = [...EXPORTACIONES_FOB_FALLBACK];
    state.exportFobSource = "fallback";
  }
  try {
    renderSerieCharts();
  } catch (error) {
    console.error("Error renderizando serie FOB, se usa vista de respaldo:", error);
    renderSerieChartsEmergencyFallback();
  }
  renderCompetidoresKpi();
  try {
    renderPropuestaTab();
  } catch (error) {
    console.error("renderPropuestaTab error:", error);
  }
  try {
    renderViabilidadTab();
  } catch (error) {
    console.error("renderViabilidadTab error:", error);
  }
}

function formatKpiTextWithBoldNumbers(value) {
  return escapeHtml(String(value ?? ""));
}

function renderKpiCardsEmergencyFallback() {
  const cards = document.getElementById("kpiCards");
  if (!cards) return;

  if (!Array.isArray(state.exportFobSeries) || !state.exportFobSeries.length) {
    state.exportFobSeries = [...EXPORTACIONES_FOB_FALLBACK];
  }
  if (!Array.isArray(state.yearsAvailable) || !state.yearsAvailable.length || !Object.keys(state.rowsByYear || {}).length) {
    seedDefaultKpiData();
  }

  const availableYears = Array.isArray(state.yearsAvailable) ? state.yearsAvailable : [];
  let selectedYear = getSelectedYear();
  if (availableYears.length && !availableYears.includes(selectedYear)) {
    selectedYear = availableYears[0];
  }
  const estado = getSelectedEstado();
  const prevYear = selectedYear - 1;

  const rowsSelectedYear = getRowsByEstado(state.rowsByYear[selectedYear] || [], estado);
  const rowsPrevYear = getRowsByEstado(state.rowsByYear[prevYear] || [], estado);
  const prodSelected = sumBy(rowsSelectedYear, "pesoKg") / 1000;
  const valSelectedUsd = mxnMillionsToUsdMillions(sumBy(rowsSelectedYear, "valorPesos") / 1_000_000);
  const prodPrev = sumBy(rowsPrevYear, "pesoKg") / 1000;
  const valPrevUsd = mxnMillionsToUsdMillions(sumBy(rowsPrevYear, "valorPesos") / 1_000_000);
  const yoyProd = prodPrev > 0 ? ((prodSelected - prodPrev) / prodPrev) * 100 : null;
  const yoyVal = valPrevUsd > 0 ? ((valSelectedUsd - valPrevUsd) / valPrevUsd) * 100 : null;
  const viajesRounded = 10500;
  const fobYear = 2025;
  const fobPrevYear = 2024;
  const fobSelected = getFobValueByYear(fobYear);
  const fobPrev = getFobValueByYear(fobPrevYear);
  const fobTrendPct =
    Number.isFinite(fobSelected) && Number.isFinite(fobPrev) && fobPrev > 0
      ? ((fobSelected - fobPrev) / fobPrev) * 100
      : null;
  const fobTrendArrow = fobTrendPct === null ? "" : fobTrendPct > 0 ? "↑" : fobTrendPct < 0 ? "↓" : "→";

  cards.innerHTML = [
    {
      title: `FOB ${fobYear} (pescados y mariscos)`,
      value: Number.isFinite(fobSelected) ? formatUsdMillionsExecutive(fobSelected) : "No disponible",
      subvalue:
        fobTrendPct === null ? "Sin tendencia disponible" : `${fobTrendArrow} ${formatPercentExecutive(fobTrendPct)} vs ${fobPrevYear}`,
      unit: "USD",
      source: "Fuente: Euromonitor Passport",
      accent: true,
    },
    {
      title: `Producción Total Nacional ${selectedYear}`,
      value: formatTonExecutive(prodSelected),
      unit: "ton",
      source: "Fuente: CONAPESCA",
    },
    {
      title: `Variación YoY ${selectedYear} vs ${prevYear}`,
      value: `Volumen: ${formatPercentExecutive(yoyProd)} | Valor: ${formatPercentExecutive(yoyVal)}`,
      unit: "%",
      source: "Fuente: CONAPESCA",
    },
    {
      title: "Viajes anuales estimados a EE.UU.",
      value: Number.isFinite(viajesRounded) ? formatNumber(viajesRounded, "viajes/año") : "No disponible",
      unit: "viajes/año",
      source: `Base: CONAPESCA 2024 (Nacional) x % exportación por especie x ${FTL_TON_POR_CAMION} ton/viaje`,
      accent: true,
    },
  ]
    .map(
      (item) => `
      <article class="kpi-card${item.accent ? " kpi-card-accent" : ""}">
        <div class="kpi-head">
          <h4>${item.title}</h4>
          <span class="kpi-unit">${item.unit}</span>
        </div>
        <div class="value">${formatKpiTextWithBoldNumbers(item.value)}</div>
        ${item.subvalue ? `<div class="kpi-subvalue">${formatKpiTextWithBoldNumbers(item.subvalue)}</div>` : ""}
        <div class="kpi-source">${item.source}</div>
      </article>
    `,
    )
    .join("");
}

function renderSerieChartsEmergencyFallback() {
  const chartEl = document.getElementById("chartSerieFobPlot");
  const notice = document.getElementById("chartSerieFobNotice");
  const insightsEl = document.getElementById("serieInsightsContent");
  const sorted = [...EXPORTACIONES_FOB_FALLBACK].sort((a, b) => a.year - b.year);
  if (!chartEl) return;

  chartEl.innerHTML = `
    <div class="serie-insight-frame">
      <strong>Vista de respaldo</strong>
      <p>Se cargó una vista alternativa de exportaciones FOB mientras se restablece la visualización principal.</p>
      <p>Último dato disponible: ${formatNumber(sorted[sorted.length - 1]?.value || 0, "USD M")} (${sorted[sorted.length - 1]?.year || "N/D"}).</p>
    </div>
  `;
  if (notice) notice.textContent = "Se mostró vista de respaldo de Exportaciones FOB.";
  if (insightsEl) {
    const sarima = calcSarimaBaseForecast(sorted, 2);
    const p2026 = sarima.points?.find((p) => p.year === 2026) || null;
    const p2027 = sarima.points?.find((p) => p.year === 2027) || null;
    const lastActual = Number(sorted[sorted.length - 1]?.value);
    const peakAnchorValue = 1130;
    const currentRefPoint = sorted.find((p) => Number(p.year) === 2024) || sorted[sorted.length - 1] || null;
    const currentRefValue = Number(currentRefPoint?.value) || 0;
    const dropFromPeakPct =
      Number.isFinite(currentRefValue) && peakAnchorValue > 0
        ? ((peakAnchorValue - currentRefValue) / peakAnchorValue) * 100
        : NaN;
    insightsEl.innerHTML = buildSerieInsightsHtml({ lastActual, p2026, p2027, dropFromPeakPct });
  }
}

function renderCompetidoresEmergencyFallback() {
  const grid = document.getElementById("competidoresGrid");
  const mapEl = document.getElementById("competidoresMap");
  if (!grid) return;
  const rowsBase = Array.isArray(competidoresData) && competidoresData.length ? competidoresData : [...competidoresFallback];
  const rows = getPanoramaCompetidores(rowsBase);
  grid.innerHTML = rows
    .map(
      (item) => `
      <article class="competidor-card">
        <h4>${item.empresa || "Competidor"}</h4>
        <div class="competidor-meta"><strong>Tipo:</strong> ${item.tipo || NO_INFO}</div>
        <div class="competidor-meta"><strong>Sede:</strong> ${item.sede || item.sedeEnMexico || item.ciudad || NO_INFO}</div>
        <div class="competidor-meta"><strong>Servicio principal:</strong> ${item.servicio || item.servicioPrincipal || NO_INFO}</div>
        <div class="competidor-meta"><strong>Sitio:</strong> ${renderWebsite(item.sitio || "")}</div>
      </article>
    `,
    )
    .join("");
  if (mapEl) mapEl.innerHTML = "";
}

function renderKpiCards() {
  const cards = document.getElementById("kpiCards");
  if (!cards) return;
  if (!Array.isArray(state.yearsAvailable) || !state.yearsAvailable.length || !Object.keys(state.rowsByYear || {}).length) {
    seedDefaultKpiData();
  }

  const availableYears = Array.isArray(state.yearsAvailable) ? state.yearsAvailable : [];
  let selectedYear = getSelectedYear();
  if (availableYears.length && !availableYears.includes(selectedYear)) {
    selectedYear = availableYears[0];
    const yearSelect = document.getElementById("filtroAnioKpi");
    if (yearSelect) yearSelect.value = String(selectedYear);
  }
  const estado = getSelectedEstado();
  const rowsSelectedYear = getRowsByEstado(state.rowsByYear[selectedYear] || [], estado);
  const rowsPrevYear = getRowsByEstado(state.rowsByYear[selectedYear - 1] || [], estado);

  const produccionTotalTon = sumBy(rowsSelectedYear, "pesoKg") / 1000;
  const valorTotalMxn = sumBy(rowsSelectedYear, "valorPesos") / 1_000_000;
  const valorTotalUsd = mxnMillionsToUsdMillions(valorTotalMxn);
  const produccionPrevTon = sumBy(rowsPrevYear, "pesoKg") / 1000;
  const valorPrevMxn = sumBy(rowsPrevYear, "valorPesos") / 1_000_000;
  const valorPrevUsd = mxnMillionsToUsdMillions(valorPrevMxn);

  const yoyProduccion = produccionPrevTon > 0 ? ((produccionTotalTon - produccionPrevTon) / produccionPrevTon) * 100 : null;
  const yoyValor = valorPrevUsd > 0 ? ((valorTotalUsd - valorPrevUsd) / valorPrevUsd) * 100 : null;
  const viajesRounded = 10500;
  const prevYear = selectedYear - 1;
  const fobYear = 2025;
  const fobPrevYear = 2024;
  const fobCurrent = getFobValueByYear(fobYear);
  const fobPrev = getFobValueByYear(fobPrevYear);
  const fobTrendPct =
    Number.isFinite(fobCurrent) && Number.isFinite(fobPrev) && fobPrev > 0
      ? ((fobCurrent - fobPrev) / fobPrev) * 100
      : null;
  const fobTrendArrow = fobTrendPct === null ? "" : fobTrendPct > 0 ? "↑" : fobTrendPct < 0 ? "↓" : "→";

  const kpis = [
    {
      title: `FOB ${fobYear} (pescados y mariscos)`,
      value: Number.isFinite(fobCurrent) ? formatUsdMillionsExecutive(fobCurrent) : "No disponible",
      subvalue: fobTrendPct === null ? "Sin tendencia disponible" : `${fobTrendArrow} ${formatPercentExecutive(fobTrendPct)} vs ${fobPrevYear}`,
      unit: "USD",
      source: "Fuente: Euromonitor Passport",
      accent: true,
    },
    {
      title: `Producción Total Nacional ${selectedYear}`,
      value: formatTonExecutive(produccionTotalTon),
      unit: "ton",
      source: "Fuente: CONAPESCA",
    },
    {
      title: `Variación YoY ${selectedYear} vs ${selectedYear - 1}`,
      value: `Volumen: ${formatPercentExecutive(yoyProduccion)} | Valor: ${formatPercentExecutive(yoyValor)}`,
      unit: "%",
      source: "Fuente: CONAPESCA",
    },
    {
      title: "Viajes anuales estimados a EE.UU.",
      value: Number.isFinite(viajesRounded) ? formatNumber(viajesRounded, "viajes/año") : "No disponible",
      unit: "viajes/año",
      source: `Base: CONAPESCA 2024 (Nacional) x % exportación por especie x ${FTL_TON_POR_CAMION} ton/viaje`,
      accent: true,
    },
  ];

  cards.innerHTML = kpis
    .map(
      (item) => `
      <article class="kpi-card${item.accent ? " kpi-card-accent" : ""}">
        <div class="kpi-head">
          <h4>${item.title}</h4>
          ${item.unit ? `<span class="kpi-unit">${item.unit}</span>` : ""}
        </div>
        <div class="value">${formatKpiTextWithBoldNumbers(item.value)}</div>
        ${item.subvalue ? `<div class="kpi-subvalue">${formatKpiTextWithBoldNumbers(item.subvalue)}</div>` : ""}
        ${item.source ? `<div class="kpi-source">${item.source}</div>` : ""}
      </article>
    `,
    )
    .join("");
}

function estimateFtlTerrestresAnuales(rows) {
  if (!Array.isArray(rows) || !rows.length) return { ftl: 0, exportTon: 0 };
  const tonBySpecies = new Map();

  rows.forEach((row) => {
    const especie = String(row?.especie || "SIN ESPECIE").trim();
    const toneladas = (Number(row?.pesoKg) || 0) / 1000;
    tonBySpecies.set(especie, (tonBySpecies.get(especie) || 0) + toneladas);
  });

  let exportTon = 0;
  tonBySpecies.forEach((toneladas, especie) => {
    exportTon += toneladas * getExportShareBySpecies(especie);
  });

  return {
    exportTon,
    ftl: exportTon / FTL_TON_POR_CAMION,
  };
}

function getExportShareBySpecies(specie) {
  const s = normalizeGeoKey(specie);

  if (s.includes("CAMAR")) return 0.35;
  if (s.includes("ATUN") || s.includes("BARRILETE") || s.includes("BONITO")) return 0.25;
  if (s.includes("PULPO") || s.includes("LANGOSTA")) return 0.3;
  if (
    s.includes("OSTION") ||
    s.includes("ALMEJA") ||
    s.includes("CARACOL") ||
    s.includes("JAIBA") ||
    s.includes("LANGOSTINO")
  ) {
    return 0.16;
  }
  if (
    s.includes("TRUCHA") ||
    s.includes("LOBINA") ||
    s.includes("ROBALO") ||
    s.includes("PARGO") ||
    s.includes("GUACHINANGO") ||
    s.includes("MERO") ||
    s.includes("CORVINA") ||
    s.includes("LENGUADO") ||
    s.includes("CABRILLA")
  ) {
    return 0.12;
  }
  if (s.includes("TIBUR") || s.includes("CAZON") || s.includes("RAYA")) return 0.08;
  if (
    s.includes("SARDINA") ||
    s.includes("ANCHOVETA") ||
    s.includes("MACARELA") ||
    s.includes("MOJARRA") ||
    s.includes("CARPA") ||
    s.includes("BAGRE") ||
    s.includes("LISA") ||
    s.includes("JUREL") ||
    s.includes("SIERRA") ||
    s.includes("BERRUGATA") ||
    s.includes("BANDERA")
  ) {
    return 0.02;
  }
  if (s.includes("OTRA")) return 0.05;
  return 0.09;
}

function getFobValueByYear(year) {
  const targetYear = Number(year);
  if (!Number.isFinite(targetYear)) return null;
  const point = (state.exportFobSeries || []).find((item) => Number(item?.year) === targetYear);
  const value = Number(point?.value);
  return Number.isFinite(value) ? value : null;
}

function initFilters() {
  const anio = document.getElementById("filtroAnioKpi");
  const estado = document.getElementById("filtroEstadoKpi");
  const mapaOrigen = document.getElementById("filtroMapaOrigen");
  const mapaTipo = document.getElementById("filtroMapaTipo");
  const mapaAnio = document.getElementById("filtroMapaAnio");

  if (!anio || !estado) return;

  const yearOptions = state.yearsAvailable.map((y) => `<option value="${y}">${y}</option>`).join("");
  anio.innerHTML = yearOptions;
  if (mapaAnio) {
    mapaAnio.innerHTML = yearOptions;
    mapaAnio.value = anio.value;
  }

  syncEstadoOptionsByYear();
  if (mapaOrigen && mapaTipo && mapaAnio) {
    syncMapaTipoOptions();
  }

  anio.onchange = () => {
    syncEstadoOptionsByYear();
    if (mapaAnio) mapaAnio.value = anio.value;
    if (mapaOrigen && mapaTipo && mapaAnio) {
      syncMapaTipoOptions();
    }
    renderKpiCards();
    if (mapaOrigen && mapaTipo && mapaAnio) {
      renderMapaEstados();
    }
  };

  estado.onchange = renderKpiCards;
  if (mapaOrigen && mapaTipo && mapaAnio) {
    mapaOrigen.onchange = () => {
      syncMapaTipoOptions();
      renderMapaEstados();
    };
    mapaTipo.onchange = renderMapaEstados;
    mapaAnio.onchange = () => {
      syncMapaTipoOptions();
      renderMapaEstados();
    };
  }
}

function syncEstadoOptionsByYear() {
  const anio = getSelectedYear();
  const estado = document.getElementById("filtroEstadoKpi");
  if (!estado) return;
  const prev = estado.value || "TODOS";
  const entidades = Array.from(new Set((state.rowsByYear[anio] || []).map((r) => r.entidad))).sort();
  estado.innerHTML = ["TODOS", ...entidades]
    .map((s) => `<option value="${s}">${titleCase(s)}</option>`)
    .join("");
  const valid = new Set(["TODOS", ...entidades]);
  estado.value = valid.has(prev) ? prev : "TODOS";
}

function syncMapaTipoOptions() {
  const tipoSelect = document.getElementById("filtroMapaTipo");
  const mapaOrigen = document.getElementById("filtroMapaOrigen");
  const mapaAnio = document.getElementById("filtroMapaAnio");
  if (!tipoSelect || !mapaOrigen || !mapaAnio) return;

  const prev = tipoSelect.value || "TODOS";
  const year = Number(mapaAnio.value || getSelectedYear());
  const origen = mapaOrigen.value || "CAPTURA";
  const especies = Array.from(
    new Set((state.rowsByYear[year] || []).filter((r) => r.origen === origen).map((r) => r.especie)),
  ).sort();

  tipoSelect.innerHTML = ["TODOS", ...especies]
    .map((s) => `<option value="${s}">${titleCase(s)}</option>`)
    .join("");

  const valid = new Set(["TODOS", ...especies]);
  tipoSelect.value = valid.has(prev) ? prev : "TODOS";
}

function renderMapaEstados() {
  const mapEl = document.getElementById("kpiEstadosMap");
  const noticeEl = document.getElementById("mapEstadosNotice");
  const mapaOrigen = document.getElementById("filtroMapaOrigen");
  const mapaTipo = document.getElementById("filtroMapaTipo");
  const mapaAnio = document.getElementById("filtroMapaAnio");
  if (!mapEl || !mapaOrigen || !mapaTipo || !mapaAnio || typeof L === "undefined") return;

  const year = Number(mapaAnio.value || getSelectedYear());
  const origen = mapaOrigen.value || "CAPTURA";
  const especie = mapaTipo.value || "TODOS";
  const rows = (state.rowsByYear[year] || []).filter(
    (r) => r.origen === origen && (especie === "TODOS" || r.especie === especie),
  );

  const byEstado = new Map();
  rows.forEach((row) => {
    byEstado.set(row.entidad, (byEstado.get(row.entidad) || 0) + row.pesoKg / 1000);
  });

  const points = Array.from(byEstado.entries())
    .map(([entidad, toneladas]) => ({ entidad, toneladas, coord: getEstadoCoord(entidad) }))
    .filter((item) => item.coord);

  if (!state.kpiEstadosMap) {
    const map = L.map(mapEl).setView([23.5, -102.5], 4.6);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "&copy; OpenStreetMap contributors",
    }).addTo(map);
    state.kpiEstadosMap = map;
  }

  if (state.kpiEstadosLayer) {
    state.kpiEstadosLayer.remove();
    state.kpiEstadosLayer = null;
  }
  const layer = L.layerGroup().addTo(state.kpiEstadosMap);
  state.kpiEstadosLayer = layer;

  if (!points.length) {
    if (noticeEl) noticeEl.textContent = "Sin datos para el filtro seleccionado.";
    state.kpiEstadosMap.setView([23.5, -102.5], 4.6);
    return;
  }

  const maxTon = Math.max(...points.map((p) => p.toneladas), 1);
  const bounds = [];
  points.forEach((p) => {
    const radius = 7 + (p.toneladas / maxTon) * 20;
    const bubbleColor = getEstadoBubbleColor(p.entidad);
    const marker = L.circleMarker([p.coord.lat, p.coord.lng], {
      radius,
      color: bubbleColor,
      fillColor: bubbleColor,
      fillOpacity: 0.45,
      weight: 1.5,
    }).addTo(layer);
    marker.bindTooltip(`${titleCase(p.entidad)}: ${formatNumber(p.toneladas, "ton")}`, {
      direction: "top",
      sticky: true,
      opacity: 0.95,
    });
    bounds.push([p.coord.lat, p.coord.lng]);
  });

  if (noticeEl) {
    noticeEl.textContent = `${titleCase(origen)} ${year} - Estados mostrados: ${points.length}.`;
  }

  state.kpiEstadosMap.fitBounds(bounds, { padding: [80, 80], maxZoom: 6 });
}

function getSelectedEstado() {
  const estado = document.getElementById("filtroEstadoKpi");
  return estado?.value || "TODOS";
}

function getSelectedYear() {
  const anio = document.getElementById("filtroAnioKpi");
  return Number(anio?.value || 2024);
}

function getRowsByEstado(rows, estado) {
  if (!estado || estado === "TODOS") return rows;
  return rows.filter((r) => r.entidad === estado);
}

function normalizeGeoKey(input) {
  return String(input || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function getEstadoCoord(entidad) {
  const key = normalizeGeoKey(entidad);
  const canonical = estadoAlias[key] || key;
  return estadoCoords[canonical] || null;
}

function getEstadoBubbleColor(entidad) {
  const key = normalizeGeoKey(entidad);
  let hash = 0;
  for (let i = 0; i < key.length; i += 1) {
    hash = (hash * 31 + key.charCodeAt(i)) | 0;
  }
  return estadoBubblePalette[Math.abs(hash) % estadoBubblePalette.length];
}

function getEmpresaMarkerColor(empresa) {
  const key = normalizeGeoKey(empresa);
  let hash = 0;
  for (let i = 0; i < key.length; i += 1) {
    hash = (hash * 31 + key.charCodeAt(i)) | 0;
  }
  return estadoBubblePalette[Math.abs(hash) % estadoBubblePalette.length];
}

function clamp(value, min, max) {
  if (!Number.isFinite(value)) return min;
  return Math.min(max, Math.max(min, value));
}

function calcStd(values) {
  const sample = (values || []).filter((v) => Number.isFinite(v));
  if (!sample.length) return 0;
  const mean = sample.reduce((acc, v) => acc + v, 0) / sample.length;
  const variance =
    sample.reduce((acc, v) => acc + (v - mean) * (v - mean), 0) / Math.max(1, sample.length - 1);
  return Math.sqrt(Math.max(variance, 0));
}

function meanTail(values, tailSize = 5) {
  const sample = (values || []).filter((v) => Number.isFinite(v));
  if (!sample.length) return 0;
  const size = Math.max(1, Math.min(sample.length, Number(tailSize) || 1));
  const tail = sample.slice(sample.length - size);
  return tail.reduce((acc, v) => acc + v, 0) / tail.length;
}

function fitAr1OnDiffs(diffs) {
  const sample = (diffs || []).filter((v) => Number.isFinite(v));
  if (sample.length < 3) return null;

  let s00 = 0;
  let s01 = 0;
  let s11 = 0;
  let sy0 = 0;
  let sy1 = 0;
  let nObs = 0;

  for (let i = 1; i < sample.length; i += 1) {
    const x0 = 1;
    const x1 = sample[i - 1];
    const y = sample[i];
    s00 += x0 * x0;
    s01 += x0 * x1;
    s11 += x1 * x1;
    sy0 += x0 * y;
    sy1 += x1 * y;
    nObs += 1;
  }

  const det = s00 * s11 - s01 * s01;
  if (!Number.isFinite(det) || Math.abs(det) < 1e-9) return null;

  const drift = (sy0 * s11 - sy1 * s01) / det;
  const phiRaw = (s00 * sy1 - s01 * sy0) / det;
  const phi = Math.max(Math.min(phiRaw, 0.98), -0.98);

  let sse = 0;
  for (let i = 1; i < sample.length; i += 1) {
    const pred = drift + phi * sample[i - 1];
    const err = sample[i] - pred;
    sse += err * err;
  }
  const sigma = Math.sqrt(sse / Math.max(1, nObs - 2));

  return {
    drift,
    phi,
    sigma: Number.isFinite(sigma) && sigma > 0 ? sigma : calcStd(sample),
    nObs,
  };
}

function calcSarimaRecentBias(series, sampleCount = 6) {
  const sorted = [...(series || [])]
    .filter((p) => Number.isFinite(p?.year) && Number.isFinite(p?.value))
    .sort((a, b) => a.year - b.year);
  if (sorted.length < 8) {
    return {
      rows: [],
      mpe: null,
      mape: null,
    };
  }

  const lookback = Math.max(3, Math.min(sorted.length - 2, Number(sampleCount) || 6));
  const startIdx = Math.max(2, sorted.length - lookback);
  const rows = [];

  for (let idx = startIdx; idx < sorted.length; idx += 1) {
    const train = sorted.slice(0, idx);
    const actual = Number(sorted[idx].value);
    if (!Number.isFinite(actual) || actual === 0) continue;
    const fc = calcSarimaBaseForecast(train, 1, { biasCorrection: false });
    const pred = Number(fc?.points?.[0]?.value);
    if (!Number.isFinite(pred)) continue;
    const err = pred - actual;
    const pe = err / actual;
    rows.push({
      year: Number(sorted[idx].year),
      pred,
      actual,
      err,
      pe,
      ape: Math.abs(pe),
    });
  }

  if (!rows.length) {
    return {
      rows: [],
      mpe: null,
      mape: null,
    };
  }

  const mpe = rows.reduce((acc, row) => acc + row.pe, 0) / rows.length;
  const mape = rows.reduce((acc, row) => acc + row.ape, 0) / rows.length;
  return { rows, mpe, mape };
}

function calcSarimaBacktestByYears(series, targetYears = [2024, 2025], options = {}) {
  const sorted = [...(series || [])]
    .filter((p) => Number.isFinite(p?.year) && Number.isFinite(p?.value))
    .sort((a, b) => a.year - b.year);
  if (!sorted.length) {
    return {
      rows: [],
      mae: null,
      mse: null,
      rmse: null,
      mape: null,
      mspe: null,
    };
  }

  const byYear = new Map(sorted.map((p) => [Number(p.year), Number(p.value)]));
  const years = Array.from(new Set((targetYears || []).map((y) => Number(y)).filter(Number.isFinite))).sort(
    (a, b) => a - b,
  );
  const rows = [];

  years.forEach((targetYear) => {
    const actual = byYear.get(targetYear);
    if (!Number.isFinite(actual) || actual === 0) return;
    const train = sorted.filter((p) => Number(p.year) < targetYear);
    if (train.length < 8) return;
    const fc = calcSarimaBaseForecast(train, 1, options);
    const pred = Number(fc?.points?.[0]?.value);
    if (!Number.isFinite(pred)) return;
    const err = pred - actual;
    const pe = err / actual;
    rows.push({
      year: targetYear,
      pred,
      actual,
      err,
      absErr: Math.abs(err),
      sqErr: err * err,
      pe,
      ape: Math.abs(pe),
      spe: pe * pe,
    });
  });

  if (!rows.length) {
    return {
      rows: [],
      mae: null,
      mse: null,
      rmse: null,
      mape: null,
      mspe: null,
    };
  }

  const mae = rows.reduce((acc, row) => acc + row.absErr, 0) / rows.length;
  const mse = rows.reduce((acc, row) => acc + row.sqErr, 0) / rows.length;
  const rmse = Math.sqrt(mse);
  const mape = rows.reduce((acc, row) => acc + row.ape, 0) / rows.length;
  const mspe = rows.reduce((acc, row) => acc + row.spe, 0) / rows.length;

  return {
    rows,
    mae,
    mse,
    rmse,
    mape,
    mspe,
  };
}

function erfApprox(x) {
  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(x);
  const t = 1 / (1 + 0.3275911 * ax);
  const y =
    1 -
    (((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) *
      t *
      Math.exp(-ax * ax));
  return sign * y;
}

function normalCdf(z) {
  if (!Number.isFinite(z)) return NaN;
  return 0.5 * (1 + erfApprox(z / Math.SQRT2));
}

function chiSquarePValueApprox(stat, dof) {
  if (!Number.isFinite(stat) || !Number.isFinite(dof) || dof <= 0) return NaN;
  if (stat <= 0) return 1;
  const z = ((stat / dof) ** (1 / 3) - (1 - 2 / (9 * dof))) / Math.sqrt(2 / (9 * dof));
  return clamp(1 - normalCdf(z), 0, 1);
}

function autocorrelationAtLag(values, lag) {
  const sample = (values || []).filter((v) => Number.isFinite(v));
  const n = sample.length;
  const k = Number(lag);
  if (n <= 1 || !Number.isFinite(k) || k < 1 || k >= n) return NaN;

  const mean = sample.reduce((acc, v) => acc + v, 0) / n;
  let num = 0;
  let den = 0;
  for (let i = 0; i < n; i += 1) {
    const centered = sample[i] - mean;
    den += centered * centered;
    if (i >= k) {
      num += centered * (sample[i - k] - mean);
    }
  }
  if (den <= 0) return NaN;
  return num / den;
}

function calcLjungBox(residuals, lag, modelDf = 1) {
  const sample = (residuals || []).filter((v) => Number.isFinite(v));
  const n = sample.length;
  const h = Number(lag);
  if (n <= 3 || !Number.isFinite(h) || h < 1 || h >= n) {
    return { lag: h, q: NaN, pValue: NaN, dof: NaN };
  }

  let qSum = 0;
  for (let k = 1; k <= h; k += 1) {
    const rk = autocorrelationAtLag(sample, k);
    if (!Number.isFinite(rk)) continue;
    qSum += (rk * rk) / (n - k);
  }
  const q = n * (n + 2) * qSum;
  const dof = Math.max(1, h - Math.max(0, Number(modelDf) || 0));
  const pValue = chiSquarePValueApprox(q, dof);
  return { lag: h, q, pValue, dof };
}

function calcResidualAutocorrelationTests(series, params = {}, options = {}) {
  const sorted = [...(series || [])]
    .filter((p) => Number.isFinite(p?.year) && Number.isFinite(p?.value))
    .sort((a, b) => a.year - b.year);
  if (sorted.length < 4) {
    return {
      residuals: [],
      nResiduals: 0,
      dw: NaN,
      lag1: NaN,
      conf95: NaN,
      ljungBox: [],
      conclusion: "Sin datos suficientes para pruebas de autocorrelación.",
    };
  }

  const values = sorted.map((p) => Number(p.value));
  const diffs = [];
  for (let i = 1; i < values.length; i += 1) {
    diffs.push(values[i] - values[i - 1]);
  }

  const drift = Number.isFinite(params?.drift) ? Number(params.drift) : meanTail(diffs, 5);
  const phi = Number.isFinite(params?.phi) ? Number(params.phi) : 0;
  const residuals = [];
  for (let i = 1; i < diffs.length; i += 1) {
    const pred = drift + phi * diffs[i - 1];
    residuals.push(diffs[i] - pred);
  }

  const n = residuals.length;
  if (n < 4) {
    return {
      residuals,
      nResiduals: n,
      dw: NaN,
      lag1: NaN,
      conf95: NaN,
      ljungBox: [],
      conclusion: "Muestra residual corta para pruebas de autocorrelación.",
    };
  }

  const maxLag = Math.max(1, Math.min(n - 1, Number(options?.maxLag) || 6));
  const modelDf = Math.max(0, Number(options?.modelDf) || 1);
  const lbLags = Array.from(new Set([1, 2, 3, 6].filter((lag) => lag <= maxLag)));
  if (!lbLags.length) lbLags.push(maxLag);
  const ljungBox = lbLags.map((lag) => calcLjungBox(residuals, lag, modelDf));

  let dwNum = 0;
  let dwDen = 0;
  for (let i = 0; i < n; i += 1) {
    const e = residuals[i];
    dwDen += e * e;
    if (i > 0) {
      const diffE = residuals[i] - residuals[i - 1];
      dwNum += diffE * diffE;
    }
  }
  const dw = dwDen > 0 ? dwNum / dwDen : NaN;
  const lag1 = autocorrelationAtLag(residuals, 1);
  const conf95 = 1.96 / Math.sqrt(n);
  const lbMain = ljungBox.find((lb) => lb.lag === 6) || ljungBox[ljungBox.length - 1];

  let conclusion = "Sin resultado concluyente.";
  if (Number.isFinite(lbMain?.pValue)) {
    if (lbMain.pValue < 0.05) {
      conclusion = "Se detecta autocorrelación residual (Ljung-Box p < 0.05).";
    } else {
      conclusion = "No se detecta autocorrelación residual significativa (Ljung-Box p >= 0.05).";
    }
  }

  return {
    residuals,
    nResiduals: n,
    dw,
    lag1,
    conf95,
    ljungBox,
    conclusion,
  };
}

function calcSarimaBaseForecast(series, horizon = 2, options = {}) {
  const sorted = [...(series || [])]
    .filter((p) => Number.isFinite(p?.year) && Number.isFinite(p?.value))
    .sort((a, b) => a.year - b.year);

  if (sorted.length < 2) {
    return {
      model: "SARIMA base (sin suficientes datos)",
      points: [],
      params: {},
    };
  }

  const years = sorted.map((p) => Number(p.year));
  const values = sorted.map((p) => Number(p.value));
  const biasCorrectionEnabled = options?.biasCorrection !== false;
  const biasSampleCount = Math.max(3, Number(options?.biasSampleCount) || 6);
  const biasStrength = Number.isFinite(options?.biasStrength) ? Number(options.biasStrength) : 0.5;
  const biasCap = Number.isFinite(options?.biasCap) ? Math.abs(Number(options.biasCap)) : 0.08;
  const diffs = [];
  for (let i = 1; i < values.length; i += 1) {
    diffs.push(values[i] - values[i - 1]);
  }

  const fit = fitAr1OnDiffs(diffs);
  const drift = Number.isFinite(fit?.drift) ? fit.drift : meanTail(diffs, 5);
  const phi = Number.isFinite(fit?.phi) ? fit.phi : 0;
  const sigma = Number.isFinite(fit?.sigma) ? fit.sigma : calcStd(diffs);
  let calibration = {
    enabled: false,
    factor: 1,
    biasPct: 0,
    mpe: null,
    mape: null,
    sampleSize: 0,
  };

  if (biasCorrectionEnabled) {
    const biasStats = calcSarimaRecentBias(sorted, biasSampleCount);
    const baseMpe = Number.isFinite(biasStats.mpe) ? biasStats.mpe : 0;
    const scaledBias = clamp(baseMpe * biasStrength, -biasCap, biasCap);
    const factor = clamp(1 - scaledBias, 0.75, 1.25);
    calibration = {
      enabled: true,
      factor,
      biasPct: scaledBias,
      mpe: biasStats.mpe,
      mape: biasStats.mape,
      sampleSize: (biasStats.rows || []).length,
    };
  }

  let prevDiff = diffs.length ? diffs[diffs.length - 1] : drift;
  let prevValue = values[values.length - 1];
  const lastYear = years[years.length - 1];
  const points = [];
  const ciZ = 1.96;

  function levelForecastStd(h, phiValue, sigmaValue) {
    const hInt = Math.max(1, Number(h) || 1);
    if (!Number.isFinite(sigmaValue) || sigmaValue <= 0) return 0;
    const phiFinite = Number.isFinite(phiValue) ? phiValue : 0;
    let coeffSqSum = 0;
    for (let m = 1; m <= hInt; m += 1) {
      const span = hInt - m + 1;
      let coeff;
      if (Math.abs(1 - phiFinite) < 1e-9) {
        coeff = span;
      } else {
        coeff = (1 - phiFinite ** span) / (1 - phiFinite);
      }
      coeffSqSum += coeff * coeff;
    }
    return Math.sqrt(coeffSqSum) * sigmaValue;
  }

  for (let h = 1; h <= Math.max(1, Number(horizon) || 1); h += 1) {
    const nextDiff = drift + phi * prevDiff;
    const nextRawValue = prevValue + nextDiff;
    const nextValue = Math.max(0, nextRawValue * calibration.factor);
    const stdH = levelForecastStd(h, phi, sigma);
    const delta = ciZ * stdH * calibration.factor;
    points.push({
      year: lastYear + h,
      value: nextValue,
      lower: Math.max(0, nextValue - delta),
      upper: nextValue + delta,
      std: stdH,
    });
    prevDiff = nextDiff;
    prevValue = nextValue;
  }

  return {
    model: fit ? "SARIMA base (ARIMA(1,1,0))" : "SARIMA base (fallback con tendencia)",
    points,
    params: { drift, phi, sigma },
    calibration,
  };
}

function toggleSerieForecastDetail(button) {
  if (!button) return;
  const wrap = button.closest(".serie-note-wrap");
  const detail = wrap ? wrap.querySelector(".serie-note-detail") : null;
  if (!detail) return;
  const showText = button.dataset.showText || "Ver detalle";
  const hideText = button.dataset.hideText || "Ocultar detalle";
  const isOpen = !detail.hasAttribute("hidden");
  if (isOpen) {
    detail.setAttribute("hidden", "");
    button.setAttribute("aria-expanded", "false");
    button.textContent = showText;
  } else {
    detail.removeAttribute("hidden");
    button.setAttribute("aria-expanded", "true");
    button.textContent = hideText;
  }
}

function getSerieScenarioLabel(scenario) {
  if (scenario === "pesimista") return "Pesimista";
  if (scenario === "optimista") return "Optimista";
  return "Base";
}

function getSerieScenarioRuleText(scenario) {
  if (scenario === "pesimista") return "pronóstico central - 1 desviación estándar";
  if (scenario === "optimista") return "pronóstico central + 1 desviación estándar";
  return "pronóstico central";
}

function setSerieScenarioButtons() {
  const buttons = document.querySelectorAll("[data-serie-scenario]");
  buttons.forEach((btn) => {
    const active = btn.dataset.serieScenario === state.serieScenario;
    btn.classList.toggle("is-active", active);
    btn.setAttribute("aria-pressed", active ? "true" : "false");
  });
}

function initSerieScenarioControls() {
  const wrap = document.getElementById("serieScenarioControls");
  if (!wrap) return;
  if (wrap.dataset.bound === "1") {
    setSerieScenarioButtons();
    return;
  }

  wrap.dataset.bound = "1";
  wrap.addEventListener("click", (event) => {
    const button = event.target.closest("[data-serie-scenario]");
    if (!button) return;
    const nextScenario = button.dataset.serieScenario || "base";
    if (!["base", "pesimista", "optimista"].includes(nextScenario)) return;
    if (state.serieScenario === nextScenario) return;
    state.serieScenario = nextScenario;
    setSerieScenarioButtons();
    renderSerieCharts();
  });
  setSerieScenarioButtons();
}

function oneSigmaFromForecastPoint(point) {
  const upper = Number(point?.upper);
  const lower = Number(point?.lower);
  const fromBounds = (upper - lower) / (2 * 1.96);
  if (Number.isFinite(fromBounds) && fromBounds > 0) return fromBounds;
  const fromStd = Number(point?.std);
  if (Number.isFinite(fromStd) && fromStd > 0) return fromStd;
  return 0;
}

function applySerieScenarioToForecast(points, scenario = "base") {
  const normalized = ["pesimista", "optimista"].includes(scenario) ? scenario : "base";
  return (points || []).map((point) => {
    const centralValue = Number(point?.value) || 0;
    const oneSigma = oneSigmaFromForecastPoint(point);
    let value = centralValue;
    if (normalized === "pesimista") value = Math.max(0, centralValue - oneSigma);
    if (normalized === "optimista") value = centralValue + oneSigma;
    return {
      ...point,
      value,
      centralValue,
      oneSigma,
      scenario: normalized,
    };
  });
}

function buildSerieAnalyticsDetail({
  scenario = "base",
  scenarioLabel = "Base",
  p2026 = null,
  p2027 = null,
  lastActual = NaN,
  baseP2026 = null,
  baseP2027 = null,
}) {
  if (!p2026 || !p2027 || !Number.isFinite(lastActual)) return "";

  const ruleText = getSerieScenarioRuleText(scenario);
  const sigma2026 = oneSigmaFromForecastPoint(baseP2026 || p2026);
  const sigma2027 = oneSigmaFromForecastPoint(baseP2027 || p2027);

  return `
    <strong>Resumen analítico</strong><br/>
    Escenario ${scenarioLabel} (${ruleText}) | Pronóstico FOB: 2026 ${formatNumber(p2026.value, "USD M")} |
    2027 ${formatNumber(p2027.value, "USD M")}.<br/><br/>
    <strong>Lógica de escenarios (+/- 1 desviación estándar)</strong><br/>
    Pesimista: pronóstico central - 1 desviación estándar.<br/>
    Base: pronóstico central.<br/>
    Optimista: pronóstico central + 1 desviación estándar.<br/><br/>
    <strong>Desviación estándar estimada</strong><br/>
    2026: ${formatNumber(sigma2026, "USD M")} | 2027: ${formatNumber(sigma2027, "USD M")}
  `;
}

function buildSerieInsightsHtml({ lastActual, p2026, p2027, dropFromPeakPct }) {
  if (!p2026 || !p2027 || !Number.isFinite(lastActual)) {
    return '<p class="serie-insight-empty">Sin datos suficientes para generar insights del forecast.</p>';
  }

  const var2026vs2025 = ((p2026.value - lastActual) / Math.max(1e-9, lastActual)) * 100;
  const var2027vs2026 = ((p2027.value - p2026.value) / Math.max(1e-9, p2026.value)) * 100;
  const dropFromPeakSigned = Number.isFinite(dropFromPeakPct) ? -Math.abs(dropFromPeakPct) : NaN;
  const dropFromPeakText = Number.isFinite(dropFromPeakSigned)
    ? formatPercent(dropFromPeakSigned)
    : "No disponible";

  return `
    <div class="serie-insight-metrics-grid">
      <div class="serie-insight-metric-card">
        <span>Variación 2026</span>
        <strong>${formatPercent(var2026vs2025)}</strong>
      </div>
      <div class="serie-insight-metric-card">
        <span>Variación 2027</span>
        <strong>${formatPercent(var2027vs2026)}</strong>
      </div>
      <div class="serie-insight-metric-card">
        <span>Desde pico 2018</span>
        <strong>${dropFromPeakText}</strong>
      </div>
    </div>
    <div class="serie-insight-frame serie-insight-neutral">
      <strong>El catalizador regulatorio</strong>
      <p>
        FSMA 204 vigente desde enero 2026. Documentación de temperatura por viaje es requisito de acceso al mercado
        estadounidense.
      </p>
    </div>
    <div class="serie-insight-frame serie-insight-cta">
      <strong>La oportunidad para CLCircular</strong>
      <p>
        Mercado bajo presión + FSMA 204 = exportadores que no pueden perder un embarque. CLCircular convierte el
        cumplimiento en un certificado automático por viaje.
      </p>
    </div>
  `;
}

function renderSerieCharts() {
  const notice = document.getElementById("chartSerieFobNotice");
  const insightsEl = document.getElementById("serieInsightsContent");
  const chartEl = document.getElementById("chartSerieFobPlot");
  setSerieScenarioButtons();
  if (!chartEl) return;
  destroyChart("chartSerieFob");
  const sorted = [...(state.exportFobSeries || [])].sort((a, b) => a.year - b.year);

  if (!sorted.length) {
    if (typeof Plotly !== "undefined") Plotly.purge(chartEl);
    if (notice) notice.textContent = "Sin datos de exportaciones FOB disponibles.";
    if (insightsEl) insightsEl.innerHTML = '<p class="serie-insight-empty">Sin datos para interpretar.</p>';
    return;
  }

  if (typeof Plotly === "undefined") {
    const sarimaNoPlotly = calcSarimaBaseForecast(sorted, 2);
    const p2026NoPlotly = sarimaNoPlotly.points?.find((p) => p.year === 2026) || null;
    const p2027NoPlotly = sarimaNoPlotly.points?.find((p) => p.year === 2027) || null;
    const lastActualNoPlotly = values[values.length - 1];
    const peakAnchorNoPlotly = 1130;
    const currentRefNoPlotly = sorted.find((p) => Number(p.year) === 2024) || sorted[sorted.length - 1] || null;
    const currentRefValueNoPlotly = Number(currentRefNoPlotly?.value) || 0;
    const dropFromPeakNoPlotly =
      Number.isFinite(currentRefValueNoPlotly) && peakAnchorNoPlotly > 0
        ? ((peakAnchorNoPlotly - currentRefValueNoPlotly) / peakAnchorNoPlotly) * 100
        : NaN;

    chartEl.innerHTML = `
      <div class="serie-insight-frame">
        <strong>Visualización temporal</strong>
        <p>No se pudo cargar Plotly. Se mantiene la lectura del forecast con respaldo interno.</p>
      </div>
    `;
    if (notice) notice.textContent = "No se pudo cargar Plotly; se mostró respaldo temporal.";
    if (insightsEl) {
      insightsEl.innerHTML = buildSerieInsightsHtml({
        lastActual: lastActualNoPlotly,
        p2026: p2026NoPlotly,
        p2027: p2027NoPlotly,
        dropFromPeakPct: dropFromPeakNoPlotly,
      });
    }
    return;
  }

  const years = sorted.map((r) => r.year);
  const values = sorted.map((r) => r.value);
  const hoverText = values.map((v) => `FOB: ${formatNumber(v, "USD M")}`);
  const sarima = calcSarimaBaseForecast(sorted, 2);
  const scenario = state.serieScenario || "base";
  const scenarioLabel = getSerieScenarioLabel(scenario);
  const baseForecastPoints = sarima.points || [];
  const forecastPoints = applySerieScenarioToForecast(baseForecastPoints, scenario);
  const forecastYears = forecastPoints.map((p) => p.year);
  const forecastValues = forecastPoints.map((p) => p.value);
  const forecastLineX = forecastPoints.length ? [years[years.length - 1], ...forecastYears] : [];
  const forecastLineY = forecastPoints.length ? [values[values.length - 1], ...forecastValues] : [];
  const xAxisMin = years[0] - 0.8;
  const xAxisMax = (forecastYears.length ? forecastYears[forecastYears.length - 1] : years[years.length - 1]) + 0.2;
  const peakAnchorValue = 1130;
  const currentRefPoint = sorted.find((p) => Number(p.year) === 2024) || sorted[sorted.length - 1] || null;
  const currentRefValue = Number(currentRefPoint?.value) || values[values.length - 1];
  const dropFromPeakPct =
    Number.isFinite(currentRefValue) && peakAnchorValue > 0
      ? ((peakAnchorValue - currentRefValue) / peakAnchorValue) * 100
      : NaN;

  const traces = [
    {
      type: "scatter",
      mode: "lines+markers",
      x: years,
      y: values,
      text: hoverText,
      hovertemplate: "Año: %{x}<br>%{text}<extra></extra>",
      line: { color: "#046f31", width: 3, shape: "spline", smoothing: 0.6 },
      marker: { color: "#046f31", size: 7 },
      fill: "tozeroy",
      fillcolor: "rgba(4, 111, 49, 0.10)",
      name: "Histórico FOB",
    },
  ];
  const annotations = [];

  if (forecastPoints.length) {
    const forecastText = [
      formatNumber(values[values.length - 1], "USD M"),
      ...forecastPoints.map((p) => formatNumber(p.value, "USD M")),
    ];
    traces.push({
      type: "scatter",
      mode: "lines+markers",
      x: forecastLineX,
      y: forecastLineY,
      text: forecastText,
      hovertemplate: "Año: %{x}<br>%{text}<extra></extra>",
      line: { color: "#0f4c81", width: 2.5, dash: "dash" },
      marker: {
        color: ["#046f31", ...forecastValues.map(() => "#0f4c81")],
        size: [6, ...forecastValues.map(() => 8)],
        symbol: ["circle", ...forecastValues.map(() => "diamond")],
      },
      name: `Pronóstico 2026-2027 (${scenarioLabel})`,
    });
  }

  Plotly.react(
    chartEl,
    traces,
    {
      margin: { l: 58, r: 18, t: 14, b: 46 },
      paper_bgcolor: "#ffffff",
      plot_bgcolor: "#ffffff",
      hovermode: "closest",
      font: { family: "Montserrat, sans-serif", color: "#1f3443", size: 10 },
      legend: {
        orientation: "h",
        x: 0,
        y: 1.08,
        font: { size: 10, color: "#4b6475" },
      },
      xaxis: {
        title: { text: "Año", font: { family: "Montserrat, sans-serif", size: 10, color: "#355264" } },
        range: [xAxisMin, xAxisMax],
        ticklabelstandoff: 8,
        automargin: true,
        tickfont: { color: "#355264", size: 9 },
        showgrid: false,
        zeroline: false,
      },
      yaxis: {
        title: {
          text: "USD millones (FOB)",
          font: { family: "Montserrat, sans-serif", size: 10, color: "#355264" },
          standoff: 18,
        },
        ticklabelstandoff: 8,
        automargin: true,
        tickfont: { color: "#355264", size: 9 },
        gridcolor: "rgba(10, 45, 74, 0.10)",
        zeroline: false,
      },
      annotations,
    },
    {
      responsive: true,
      displaylogo: false,
      scrollZoom: false,
      modeBarButtonsToRemove: ["lasso2d", "select2d", "autoScale2d", "toggleSpikelines"],
    },
  );
  if (notice) {
    const p2026 = forecastPoints.find((p) => p.year === 2026);
    const p2027 = forecastPoints.find((p) => p.year === 2027);
    if (p2026 && p2027) {
      const lastActual = values[values.length - 1];
      const baseP2026 = baseForecastPoints.find((p) => p.year === 2026) || null;
      const baseP2027 = baseForecastPoints.find((p) => p.year === 2027) || null;
      const detailHtml = buildSerieAnalyticsDetail({
        scenario,
        scenarioLabel,
        p2026,
        p2027,
        lastActual,
        baseP2026,
        baseP2027,
      });
      notice.innerHTML = `
        <span class="serie-note-wrap">
          <button
            type="button"
            class="competidor-btn serie-note-btn"
            aria-expanded="false"
            data-show-text="Ver info analítica"
            data-hide-text="Ocultar info analítica"
            onclick="toggleSerieForecastDetail(this)"
          >
            Ver info analítica
          </button>
          <span class="serie-note-detail" hidden>${detailHtml}</span>
        </span>
      `;
    } else {
      notice.textContent = `${sarima.model} sin variables exógenas.`;
    }
  }

  if (insightsEl) {
    const p2026Insight = forecastPoints.find((p) => p.year === 2026);
    const p2027Insight = forecastPoints.find((p) => p.year === 2027);
    const lastActual = values[values.length - 1];
    insightsEl.innerHTML = buildSerieInsightsHtml({
      lastActual,
      p2026: p2026Insight,
      p2027: p2027Insight,
      dropFromPeakPct,
    });
  }
}

function aggregateTop(rows, key, n) {
  const acc = new Map();
  for (const row of rows) {
    const k = row[key] || "SIN DATO";
    acc.set(k, (acc.get(k) || 0) + row.pesoKg / 1000);
  }
  return Array.from(acc.entries())
    .map(([label, value]) => ({ label: titleCase(label), value }))
    .sort((a, b) => b.value - a.value)
    .slice(0, n);
}

function renderBarChart(canvasId, data, color, unit = "ton") {
  destroyChart(canvasId);
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;

  state.charts[canvasId] = new Chart(ctx, {
    type: "bar",
    data: {
      labels: data.map((d) => d.label),
      datasets: [
        {
          data: data.map((d) => d.value),
          backgroundColor: color,
          borderRadius: 8,
        },
      ],
    },
    options: {
      animation: { duration: 700 },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (context) => `${formatNumber(context.parsed.y, unit)}`,
          },
        },
      },
      scales: {
        x: {
          ticks: { color: "#284257", maxRotation: 35, minRotation: 20 },
          grid: { display: false },
        },
        y: {
          ticks: {
            color: "#36556b",
            callback: (v) => `${Number(v).toLocaleString("es-MX")}`,
          },
          grid: { color: "rgba(10, 45, 74, 0.08)" },
        },
      },
    },
  });
}

function renderLineChart(canvasId, labels, values, color, unit = "") {
  destroyChart(canvasId);
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;

  state.charts[canvasId] = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          data: values,
          borderColor: color,
          backgroundColor: `${color}33`,
          pointRadius: 3,
          pointHoverRadius: 6,
          pointHitRadius: 16,
          fill: true,
          tension: 0.3,
        },
      ],
    },
    options: {
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          enabled: true,
          callbacks: {
            title: (items) => `Año: ${items[0]?.label || ""}`,
            label: (context) => (unit ? `Valor: ${formatNumber(context.parsed.y, unit)}` : formatNumber(context.parsed.y)),
          },
        },
      },
      scales: {
        x: { ticks: { color: "#284257" }, grid: { display: false } },
        y: {
          ticks: {
            color: "#36556b",
            callback: (v) => Number(v).toLocaleString("es-MX"),
          },
          grid: { color: "rgba(10, 45, 74, 0.08)" },
        },
      },
    },
  });
}

function destroyChart(id) {
  if (state.charts[id]) {
    state.charts[id].destroy();
  }
}

function resizeAllCharts() {
  Object.values(state.charts).forEach((chart) => chart.resize());
  const seriePlot = document.getElementById("chartSerieFobPlot");
  if (seriePlot && typeof Plotly !== "undefined" && seriePlot.data) {
    Plotly.Plots.resize(seriePlot);
  }
  if (typeof Plotly !== "undefined") {
    [
      "viabChartRevenue",
      "viabChartFcf",
      "viabChartAmort",
      "viabChartFiscal",
      "viabChartScenario",
      "clusterScatterPlot",
      "clusterPcaPlot",
      "clusterElbowPlot",
    ].forEach((id) => {
      const chartEl = document.getElementById(id);
      if (chartEl && chartEl.data) Plotly.Plots.resize(chartEl);
    });
  }
}

function sumBy(arr, field) {
  return arr.reduce((acc, x) => acc + (Number(x[field]) || 0), 0);
}

function formatNumber(value, suffix = "") {
  const formatted = Number(value).toLocaleString("es-MX", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  });
  return suffix ? `${formatted} ${suffix}` : formatted;
}

function formatTonExecutive(valueTon) {
  const value = Number(valueTon) || 0;
  const abs = Math.abs(value);
  if (abs >= 1_000_000) {
    return `${(value / 1_000_000).toLocaleString("es-MX", {
      minimumFractionDigits: 1,
      maximumFractionDigits: 1,
    })}M ton`;
  }
  if (abs >= 1_000) {
    return `${(value / 1_000).toLocaleString("es-MX", {
      minimumFractionDigits: 1,
      maximumFractionDigits: 1,
    })}k ton`;
  }
  return `${Math.round(value).toLocaleString("es-MX")} ton`;
}

function formatMxnMillionsExecutive(valueMillionsMxn) {
  const value = Number(valueMillionsMxn) || 0;
  const abs = Math.abs(value);
  if (abs >= 1_000) {
    return `${(value / 1_000).toLocaleString("es-MX", {
      minimumFractionDigits: 1,
      maximumFractionDigits: 1,
    })}B MXN`;
  }
  return `${Math.round(value).toLocaleString("es-MX")}M MXN`;
}

function formatMxnExecutive(valueMxn) {
  const value = Number(valueMxn) || 0;
  const abs = Math.abs(value);
  if (abs >= 1_000_000_000) {
    return `${(value / 1_000_000_000).toLocaleString("es-MX", {
      minimumFractionDigits: 1,
      maximumFractionDigits: 1,
    })}B MXN`;
  }
  if (abs >= 1_000_000) {
    return `${(value / 1_000_000).toLocaleString("es-MX", {
      minimumFractionDigits: 1,
      maximumFractionDigits: 1,
    })}M MXN`;
  }
  return `${Math.round(value).toLocaleString("es-MX")} MXN`;
}

function mxnMillionsToUsdMillions(valueMillionsMxn, fxRate = KPI_USD_MXN_RATE) {
  const value = Number(valueMillionsMxn);
  const fx = Number(fxRate);
  if (!Number.isFinite(value) || !Number.isFinite(fx) || fx <= 0) return 0;
  return value / fx;
}

function formatUsdMillionsExecutive(valueUsdMillions) {
  const value = Number(valueUsdMillions) || 0;
  return `$${value.toLocaleString("es-MX", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  })}M USD`;
}

function formatCurrencyMillions(value) {
  return `${Number(value).toLocaleString("es-MX", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })} M MXN`;
}

function formatCurrency(value) {
  return Number(value).toLocaleString("es-MX", {
    style: "currency",
    currency: "MXN",
    maximumFractionDigits: 0,
  });
}

function formatPercent(value) {
  if (value === null || Number.isNaN(value)) return "No disponible";
  const sign = value > 0 ? "+" : "";
  return `${sign}${Number(value).toLocaleString("es-MX", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}%`;
}

function formatPercentExecutive(value) {
  if (value === null || Number.isNaN(value)) return "N/D";
  const sign = value > 0 ? "+" : "";
  return `${sign}${Number(value).toLocaleString("es-MX", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  })}%`;
}

function formatPercentUnsigned(value) {
  if (value === null || Number.isNaN(value)) return "No disponible";
  return `${Number(value).toLocaleString("es-MX", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}%`;
}

function titleCase(input) {
  return String(input).replace(/\p{L}+/gu, (word) => {
    const lower = word.toLocaleLowerCase("es-MX");
    return lower.charAt(0).toLocaleUpperCase("es-MX") + lower.slice(1);
  });
}

function setStatus(message, ok) {
  const el = document.getElementById("fileStatus");
  if (el) {
    el.textContent = message;
    el.style.borderColor = ok ? "#8dd7b6" : "#f0b7b7";
    el.style.background = ok ? "#edfff6" : "#fff4f4";
    return;
  }

  const uploadLabel = document.querySelector("label.upload-btn[for='xlsxInput']");
  if (!uploadLabel) return;
  const original = uploadLabel.dataset.originalText || uploadLabel.textContent || "Cargar XLSX manual";
  uploadLabel.dataset.originalText = original;
  uploadLabel.textContent = ok ? "XLSX cargado ✓" : "Error al cargar XLSX";
  window.setTimeout(() => {
    const currentLabel = document.querySelector("label.upload-btn[for='xlsxInput']");
    if (!currentLabel) return;
    currentLabel.textContent = currentLabel.dataset.originalText || "Cargar XLSX manual";
  }, 2400);
}
