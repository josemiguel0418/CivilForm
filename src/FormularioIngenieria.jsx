import { useState } from "react";
import * as XLSX from "xlsx";
import { MapPin, Building2, Layers, Calculator, ChevronRight, ChevronLeft, AlertTriangle, UploadCloud, CheckCircle2 } from "lucide-react";

// ---------------------------------------------------------------------------
// CARGA DESDE EXCEL
// El Excel que suba el ingeniero debe tener una fila de datos con columnas
// EXACTAMENTE con estos nombres de encabezado (no importa el orden de columnas,
// sí importa que el texto del encabezado coincida):
//
// proyecto | municipio | zona | suelo | capacidad_portante | profundidad_desplante
// | uso | pisos | area_por_piso | altura_entrepiso | sistema | fc | fy | carga_viva | carga_muerta
//
// Valores válidos para columnas de tipo "selección":
//   zona: baja | intermedia | alta
//   suelo: A | B | C | D | E
//   uso: vivienda | oficina | comercial | educativo | hospitalario | industrial
// ---------------------------------------------------------------------------
const MAPA_COLUMNAS = {
  proyecto: "proyecto",
  municipio: "municipio",
  zona: "zona",
  suelo: "suelo",
  capacidadPortante: "capacidad_portante",
  profundidadDesplante: "profundidad_desplante",
  uso: "uso",
  pisos: "pisos",
  areaPorPiso: "area_por_piso",
  alturaEntrepiso: "altura_entrepiso",
  sistema: "sistema",
  fc: "fc",
  fy: "fy",
  cargaViva: "carga_viva",
  cargaMuerta: "carga_muerta",
};

// ---------------------------------------------------------------------------
// TABLAS DE REFERENCIA (EJEMPLO — reemplazar con los valores reales que te dé
// el ingeniero civil, citando la norma exacta, p. ej. NSR-10 Título A y B)
// ---------------------------------------------------------------------------
const TIPOS_SUELO = [
  { id: "A", label: "A — Roca competente" },
  { id: "B", label: "B — Roca de rigidez media" },
  { id: "C", label: "C — Suelo denso a firme" },
  { id: "D", label: "D — Suelo medianamente denso" },
  { id: "E", label: "E — Suelo blando" },
];

const ZONAS_SISMICAS = [
  { id: "baja", label: "Baja amenaza sísmica", coefEjemplo: 0.10 },
  { id: "intermedia", label: "Amenaza sísmica intermedia", coefEjemplo: 0.20 },
  { id: "alta", label: "Amenaza sísmica alta", coefEjemplo: 0.35 },
];

const TIPOS_USO = [
  { id: "vivienda", label: "Vivienda", grupoUso: "I", coefImportancia: 1.0, cargaVivaEjemplo: 1.8 },
  { id: "oficina", label: "Edificio de oficinas", grupoUso: "I", coefImportancia: 1.0, cargaVivaEjemplo: 2.0 },
  { id: "comercial", label: "Comercial", grupoUso: "I", coefImportancia: 1.0, cargaVivaEjemplo: 3.0 },
  { id: "educativo", label: "Institución educativa", grupoUso: "II", coefImportancia: 1.1, cargaVivaEjemplo: 2.0 },
  { id: "hospitalario", label: "Hospitalario", grupoUso: "IV", coefImportancia: 1.5, cargaVivaEjemplo: 3.0 },
  { id: "industrial", label: "Industrial", grupoUso: "I", coefImportancia: 1.0, cargaVivaEjemplo: 5.0 },
];

const SISTEMAS_ESTRUCTURALES = [
  "Pórticos de concreto reforzado",
  "Muros de carga en concreto",
  "Mampostería confinada",
  "Estructura metálica",
  "Estructura en madera",
];

const STEPS = [
  { id: 0, label: "Ubicación y suelo", icon: MapPin },
  { id: 1, label: "Uso y geometría", icon: Building2 },
  { id: 2, label: "Materiales y cargas", icon: Layers },
  { id: 3, label: "Resultados", icon: Calculator },
];

const inputBase =
  "w-full rounded-md border border-[#CBD3DC] bg-white px-3 py-2 text-[15px] text-[#1B2430] outline-none transition-colors focus:border-[#C4571B] focus:ring-2 focus:ring-[#C4571B]/20";
const labelBase = "mb-1.5 block text-[13px] font-medium text-[#4B5566]";
const unitTag = "ml-2 font-mono text-[12px] text-[#8A94A3]";

function Field({ label, unit, children }) {
  return (
    <div>
      <label className={labelBase}>
        {label}
        {unit && <span className={unitTag}>{unit}</span>}
      </label>
      {children}
    </div>
  );
}

export default function FormularioIngenieria() {
  const [step, setStep] = useState(0);
  const [data, setData] = useState({
    proyecto: "",
    municipio: "",
    zona: "intermedia",
    suelo: "C",
    capacidadPortante: "150",
    profundidadDesplante: "1.5",

    uso: "vivienda",
    pisos: "3",
    areaPorPiso: "80",
    alturaEntrepiso: "2.6",
    sistema: SISTEMAS_ESTRUCTURALES[0],

    fc: "21",
    fy: "420",
    cargaViva: "1.8",
    cargaMuerta: "3.5",
  });

  const [cargaInfo, setCargaInfo] = useState(null); // { tipo: 'ok'|'error', mensaje }

  const set = (key) => (e) => setData((d) => ({ ...d, [key]: e.target.value }));

  const usoInfo = TIPOS_USO.find((u) => u.id === data.uso);
  const zonaInfo = ZONAS_SISMICAS.find((z) => z.id === data.zona);

  // -------------------------------------------------------------------------
  // Lee el primer archivo Excel/CSV subido, toma la PRIMERA fila de datos,
  // y llena el formulario buscando cada columna por su nombre de encabezado
  // (ver MAPA_COLUMNAS arriba). Las columnas que no encuentre simplemente
  // no se tocan, quedan con el valor que ya tenía el formulario.
  // -------------------------------------------------------------------------
  function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const hoja = wb.Sheets[wb.SheetNames[0]];
        const filas = XLSX.utils.sheet_to_json(hoja, { defval: "" });

        if (!filas.length) {
          setCargaInfo({ tipo: "error", mensaje: "El archivo no tiene filas de datos debajo del encabezado." });
          return;
        }

        const fila = filas[0]; // toma la primera fila de datos
        const encontradas = [];
        const faltantes = [];

        setData((d) => {
          const nuevo = { ...d };
          for (const [campoFormulario, nombreColumna] of Object.entries(MAPA_COLUMNAS)) {
            const valor = fila[nombreColumna];
            if (valor !== undefined && valor !== "") {
              nuevo[campoFormulario] = String(valor);
              encontradas.push(nombreColumna);
            } else {
              faltantes.push(nombreColumna);
            }
          }
          return nuevo;
        });

        setCargaInfo({
          tipo: "ok",
          mensaje: `Se llenaron ${encontradas.length} de ${Object.keys(MAPA_COLUMNAS).length} campos. ${
            faltantes.length ? `No se encontraron las columnas: ${faltantes.join(", ")}.` : "Todas las columnas se encontraron."
          }`,
        });
      } catch (err) {
        setCargaInfo({ tipo: "error", mensaje: "No se pudo leer el archivo. ¿Es un .xlsx o .csv válido?" });
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // -------------------------------------------------------------------------
  // MOTOR DE CÁLCULO — ILUSTRATIVO ÚNICAMENTE.
  // Estas fórmulas NO son las de una norma real; existen solo para mostrar
  // dónde y cómo conectar la lógica una vez el ingeniero te entregue las
  // fórmulas correctas y su referencia normativa exacta.
  // -------------------------------------------------------------------------
  function calcular() {
    const pisos = parseFloat(data.pisos) || 0;
    const area = parseFloat(data.areaPorPiso) || 0;
    const cv = parseFloat(data.cargaViva) || 0;
    const cm = parseFloat(data.cargaMuerta) || 0;

    const areaTotal = pisos * area;
    const pesoSismico = areaTotal * (cm + 0.25 * cv); // ejemplo simplificado
    const coefSismicoBase = zonaInfo?.coefEjemplo ?? 0;
    const coefAjustado = coefSismicoBase * (usoInfo?.coefImportancia ?? 1);
    const cortanteBasal = pesoSismico * coefAjustado;

    return { areaTotal, pesoSismico, coefAjustado, cortanteBasal };
  }

  const resultados = calcular();

  return (
    <div className="min-h-[600px] w-full bg-[#E9EDF1] p-6 font-sans text-[#1B2430]">
      <div className="mx-auto max-w-3xl">
        <header className="mb-6">
          <p className="font-mono text-[12px] uppercase tracking-wide text-[#8A94A3]">
            Ficha técnica de proyecto
          </p>
          <h1 className="mt-1 text-[26px] font-semibold leading-tight">
            Datos de entrada para diseño estructural
          </h1>
        </header>

        {/* Carga de datos desde Excel */}
        <div className="mb-6 rounded-lg border border-[#CBD3DC] bg-white p-4">
          <div className="flex flex-wrap items-center gap-3">
            <label className="flex cursor-pointer items-center gap-2 rounded-md border border-[#C4571B] bg-[#F3E4D8] px-3 py-2 text-[13px] font-medium text-[#8A4B14]">
              <UploadCloud size={16} />
              Cargar datos desde Excel/CSV
              <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleExcelUpload} />
            </label>
            <p className="text-[12px] text-[#8A94A3]">
              El archivo debe tener columnas con encabezados: proyecto, municipio, zona, suelo, capacidad_portante,
              profundidad_desplante, uso, pisos, area_por_piso, altura_entrepiso, sistema, fc, fy, carga_viva, carga_muerta.
            </p>
          </div>
          {cargaInfo && (
            <div
              className={`mt-3 flex gap-2 rounded-md p-2 text-[13px] ${
                cargaInfo.tipo === "ok" ? "bg-[#E7F3ED] text-[#2F6B4F]" : "bg-[#FBE9E7] text-[#B3401E]"
              }`}
            >
              {cargaInfo.tipo === "ok" ? <CheckCircle2 size={16} className="mt-0.5 shrink-0" /> : <AlertTriangle size={16} className="mt-0.5 shrink-0" />}
              <p>{cargaInfo.mensaje}</p>
            </div>
          )}
        </div>

        {/* Rail de pasos */}
        <div className="mb-6 flex items-stretch overflow-hidden rounded-lg border border-[#CBD3DC] bg-white">
          {STEPS.map((s, i) => {
            const Icon = s.icon;
            const active = step === s.id;
            const done = step > s.id;
            return (
              <button
                key={s.id}
                onClick={() => setStep(s.id)}
                className={`flex flex-1 items-center gap-2 border-r border-[#E4E8ED] px-3 py-3 text-left text-[13px] transition-colors last:border-r-0 ${
                  active ? "bg-[#1B2430] text-white" : done ? "bg-[#F3E4D8] text-[#1B2430]" : "text-[#8A94A3] hover:bg-[#F5F7F9]"
                }`}
              >
                <Icon size={16} className="shrink-0" />
                <span className="hidden sm:inline">
                  {String(s.id + 1).padStart(2, "0")} — {s.label}
                </span>
              </button>
            );
          })}
        </div>

        <div className="rounded-lg border border-[#CBD3DC] bg-white p-6">
          {step === 0 && (
            <div className="grid grid-cols-2 gap-4">
              <div className="col-span-2">
                <Field label="Nombre del proyecto">
                  <input className={inputBase} value={data.proyecto} onChange={set("proyecto")} placeholder="Ej. Edificio Los Almendros" />
                </Field>
              </div>
              <Field label="Municipio / ciudad">
                <input className={inputBase} value={data.municipio} onChange={set("municipio")} placeholder="Ej. Itagüí" />
              </Field>
              <Field label="Zona de amenaza sísmica">
                <select className={inputBase} value={data.zona} onChange={set("zona")}>
                  {ZONAS_SISMICAS.map((z) => (
                    <option key={z.id} value={z.id}>{z.label}</option>
                  ))}
                </select>
              </Field>
              <Field label="Tipo de suelo">
                <select className={inputBase} value={data.suelo} onChange={set("suelo")}>
                  {TIPOS_SUELO.map((s) => (
                    <option key={s.id} value={s.id}>{s.label}</option>
                  ))}
                </select>
              </Field>
              <Field label="Capacidad portante" unit="kPa">
                <input type="number" className={inputBase} value={data.capacidadPortante} onChange={set("capacidadPortante")} />
              </Field>
              <Field label="Profundidad de desplante" unit="m">
                <input type="number" className={inputBase} value={data.profundidadDesplante} onChange={set("profundidadDesplante")} />
              </Field>
            </div>
          )}

          {step === 1 && (
            <div className="grid grid-cols-2 gap-4">
              <div className="col-span-2">
                <Field label="Tipo de uso / ocupación">
                  <select className={inputBase} value={data.uso} onChange={set("uso")}>
                    {TIPOS_USO.map((u) => (
                      <option key={u.id} value={u.id}>{u.label}</option>
                    ))}
                  </select>
                </Field>
              </div>
              <div className="col-span-2 rounded-md bg-[#F5F7F9] px-3 py-2 text-[13px] text-[#4B5566]">
                Grupo de uso <span className="font-mono">{usoInfo?.grupoUso}</span> · Coeficiente de importancia{" "}
                <span className="font-mono">{usoInfo?.coefImportancia}</span> — se aplica automáticamente en resultados.
              </div>
              <Field label="Número de pisos">
                <input type="number" className={inputBase} value={data.pisos} onChange={set("pisos")} />
              </Field>
              <Field label="Área por piso" unit="m²">
                <input type="number" className={inputBase} value={data.areaPorPiso} onChange={set("areaPorPiso")} />
              </Field>
              <Field label="Altura de entrepiso" unit="m">
                <input type="number" className={inputBase} value={data.alturaEntrepiso} onChange={set("alturaEntrepiso")} />
              </Field>
              <Field label="Sistema estructural">
                <select className={inputBase} value={data.sistema} onChange={set("sistema")}>
                  {SISTEMAS_ESTRUCTURALES.map((s) => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
              </Field>
            </div>
          )}

          {step === 2 && (
            <div className="grid grid-cols-2 gap-4">
              <Field label="f'c del concreto" unit="MPa">
                <input type="number" className={inputBase} value={data.fc} onChange={set("fc")} />
              </Field>
              <Field label="fy del acero" unit="MPa">
                <input type="number" className={inputBase} value={data.fy} onChange={set("fy")} />
              </Field>
              <Field label="Carga viva" unit="kN/m²">
                <input type="number" className={inputBase} value={data.cargaViva} onChange={set("cargaViva")} />
              </Field>
              <Field label="Carga muerta adicional" unit="kN/m²">
                <input type="number" className={inputBase} value={data.cargaMuerta} onChange={set("cargaMuerta")} />
              </Field>
              <div className="col-span-2 text-[13px] text-[#8A94A3]">
                Sugerencia según uso «{usoInfo?.label}»: carga viva típica ≈ {usoInfo?.cargaVivaEjemplo} kN/m² (valor de referencia, confírmalo con el ingeniero).
              </div>
            </div>
          )}

          {step === 3 && (
            <div>
              <div className="mb-4 flex gap-2 rounded-md border border-[#E9C79A] bg-[#FBF1E6] p-3 text-[13px] text-[#8A4B14]">
                <AlertTriangle size={18} className="mt-0.5 shrink-0" />
                <p>
                  Las fórmulas usadas aquí son solo de ejemplo, para mostrar dónde conectar el cálculo. No están tomadas
                  de una norma real. Reemplaza la función <code className="font-mono">calcular()</code> del código con
                  las fórmulas exactas que te entregue el ingeniero, citando la norma y el artículo correspondiente.
                </p>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <ResultCard label="Área construida total" value={resultados.areaTotal.toFixed(1)} unit="m²" />
                <ResultCard label="Peso sísmico estimado (W)" value={resultados.pesoSismico.toFixed(1)} unit="kN" />
                <ResultCard label="Coeficiente sísmico ajustado" value={resultados.coefAjustado.toFixed(3)} unit="" />
                <ResultCard label="Cortante basal estimado (V)" value={resultados.cortanteBasal.toFixed(1)} unit="kN" highlight />
              </div>
            </div>
          )}

          <div className="mt-6 flex justify-between border-t border-[#E4E8ED] pt-4">
            <button
              onClick={() => setStep((s) => Math.max(0, s - 1))}
              disabled={step === 0}
              className="flex items-center gap-1 rounded-md px-3 py-2 text-[14px] text-[#4B5566] disabled:opacity-30"
            >
              <ChevronLeft size={16} /> Anterior
            </button>
            <button
              onClick={() => setStep((s) => Math.min(STEPS.length - 1, s + 1))}
              disabled={step === STEPS.length - 1}
              className="flex items-center gap-1 rounded-md bg-[#1B2430] px-4 py-2 text-[14px] text-white disabled:opacity-30"
            >
              Siguiente <ChevronRight size={16} />
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function ResultCard({ label, value, unit, highlight }) {
  return (
    <div className={`rounded-md border p-4 ${highlight ? "border-[#C4571B] bg-[#F3E4D8]" : "border-[#E4E8ED] bg-[#F5F7F9]"}`}>
      <p className="text-[13px] text-[#4B5566]">{label}</p>
      <p className="mt-1 font-mono text-[22px] font-semibold text-[#1B2430]">
        {value} <span className="text-[14px] text-[#8A94A3]">{unit}</span>
      </p>
    </div>
  );
}
