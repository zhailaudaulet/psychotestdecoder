import React, { useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import { Upload, Download, FileSpreadsheet, CheckCircle2, AlertCircle, X } from "lucide-react";
import "./styles.css";

const ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
const MODULUS = 10 ** 10;
const OFFSET = 3089424140;

const ORDER_LOW_TO_HIGH = [
  "Шкала искренности", "Лабильность", "Тревожность", "Сензитивность",
  "Интроверсия", "Ригидность", "Агрессивность", "Спонтанность",
  "Экстраверсия", "Шкала аггравации",
];

const OUT_ORDER = [
  "Экстраверсия", "Спонтанность", "Агрессивность", "Ригидность",
  "Интроверсия", "Сензитивность", "Тревожность", "Лабильность",
  "Шкала искренности", "Шкала аггравации",
];

function charValue(c) {
  const v = ALPHABET.indexOf(c);
  if (v === -1) throw new Error(`Недопустимый символ в коде: ${c}`);
  return v >= 52 ? v - 64 : v;
}

function decode(payload) {
  let n = 0;
  for (const c of payload) n = n * 64 + charValue(c);
  n = (n + OFFSET) % MODULUS;
  const digits = String(n).padStart(10, "0").split("").reverse().map(Number);
  return Object.fromEntries(ORDER_LOW_TO_HIGH.map((name, i) => [name, digits[i]]));
}

function level(score) {
  if (score <= 1) return "гипоэмотивность";
  if (score <= 4) return "норма";
  if (score <= 7) return "акцентуация";
  return "избыточность";
}

function decodeUrl(url) {
  const match = String(url ?? "").match(/[?&]v=([A-Za-z0-9_-]+)/);
  if (!match) throw new Error("Не найдена ссылка с параметром v");

  const payload = match[1];
  if (!payload.startsWith("ito") || payload.length !== 10) {
    throw new Error("Неожиданный формат ссылки");
  }

  const code = payload.slice(3);
  if (code[0] !== "A") {
    throw new Error(`Версия/тип теста не подтверждён: первый символ ${code[0]}`);
  }

  return decode(code.slice(1));
}

const OUTPUT_HEADERS = OUT_ORDER.flatMap((name) => [name, "уровень"]);

function buildResult(name, url) {
  try {
    const scores = decodeUrl(url);
    const row = { "ФИО": name ?? "" };

    for (const metric of OUT_ORDER) {
      row[metric] = scores[metric];
      if (metric !== "Шкала искренности" && metric !== "Шкала аггравации") {
        row["уровень_" + metric] = level(scores[metric]);
      } else {
        row["уровень_" + metric] = level(scores[metric]);
      }
    }

    return { ok: true, row };
  } catch (error) {
    return { ok: false, name: name ?? "", error: error.message };
  }
}

function exportExcel(rows, errors) {
  const data = rows.map((r) => {
    const out = { "ФИО": r["ФИО"] };
    for (const metric of OUT_ORDER) {
      out[metric] = r[metric];
      out["уровень"] = undefined;
      // Temporary placeholder is removed below; each metric gets its own level column.
    }
    return out;
  });

  const normalized = rows.map((r) => {
    const out = { "ФИО": r["ФИО"] };
    for (const metric of OUT_ORDER) {
      out[metric] = r[metric];
      out[metric + "_уровень"] = r["уровень_" + metric];
    }
    return out;
  });

  // Exact requested column order: metric, уровень, metric, уровень...
  const finalRows = normalized.map((r) => {
    const out = { "ФИО": r["ФИО"] };
    for (const metric of OUT_ORDER) {
      out[metric] = r[metric];
      out["уровень"] = r[metric + "_уровень"];
      // XLSX cannot have duplicate header names reliably, so duplicate level headers
      // are generated later with explicit worksheet header rows.
    }
    return out;
  });

  const header = ["ФИО", ...OUTPUT_HEADERS];
  const matrix = [header];

  for (const r of rows) {
    const line = [r["ФИО"]];
    for (const metric of OUT_ORDER) {
      line.push(r[metric], r["уровень_" + metric]);
    }
    matrix.push(line);
  }

  // Add failed rows as a separate sheet so no input row silently disappears.
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(matrix);
  ws["!cols"] = [{ wch: 30 }, ...Array(OUTPUT_HEADERS.length).fill({ wch: 18 })];
  XLSX.utils.book_append_sheet(wb, ws, "Результаты");

  if (errors.length) {
    const errorMatrix = [["ФИО", "Ссылка", "Ошибка"], ...errors.map(e => [e.name, e.url, e.error])];
    const ew = XLSX.utils.aoa_to_sheet(errorMatrix);
    ew["!cols"] = [{ wch: 30 }, { wch: 60 }, { wch: 45 }];
    XLSX.utils.book_append_sheet(wb, ew, "Ошибки");
  }

  XLSX.writeFile(wb, "psytest_results.xlsx");
}

function App() {
  const [file, setFile] = useState(null);
  const [results, setResults] = useState([]);
  const [errors, setErrors] = useState([]);
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);

  const stats = useMemo(() => ({
    total: results.length + errors.length,
    success: results.length,
    failed: errors.length
  }), [results, errors]);

  async function processFile(selectedFile) {
    if (!selectedFile) return;
    if (!/\.(xlsx|xls)$/i.test(selectedFile.name)) {
      alert("Загрузите Excel-файл .xlsx или .xls");
      return;
    }

    setFile(selectedFile);
    setLoading(true);
    setResults([]);
    setErrors([]);

    try {
      const buffer = await selectedFile.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

      if (!data.length) throw new Error("Excel-файл пустой.");

      const keys = Object.keys(data[0]);
      const nameKey = keys.find(k => String(k).trim().toLowerCase() === "фио");
      const urlKey = keys.find(k => String(k).trim().toLowerCase() === "ссылка");

      if (!nameKey || !urlKey) {
        throw new Error('В файле должны быть столбцы "ФИО" и "ссылка".');
      }

      const good = [];
      const bad = [];

      for (const item of data) {
        const name = String(item[nameKey] ?? "").trim();
        const url = String(item[urlKey] ?? "").trim();

        if (!name && !url) continue;

        const result = buildResult(name, url);
        if (result.ok) good.push(result.row);
        else bad.push({ name, url, error: result.error });
      }

      setResults(good);
      setErrors(bad);
    } catch (e) {
      alert(e.message || "Не удалось обработать файл.");
      setFile(null);
    } finally {
      setLoading(false);
    }
  }

  function reset() {
    setFile(null);
    setResults([]);
    setErrors([]);
  }

  return (
    <div className="page">
      <div className="container">
        <header className="header">
          <div className="brand-mark"><FileSpreadsheet size={22} /></div>
          <div>
            <h1>PsychTest Decoder</h1>
            <p>Массовая расшифровка результатов из Excel</p>
          </div>
        </header>

        <main>
          {!file && (
            <section
              className={`dropzone ${dragging ? "dragging" : ""}`}
              onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragging(false);
                processFile(e.dataTransfer.files?.[0]);
              }}
              onClick={() => document.getElementById("file-input").click()}
            >
              <input
                id="file-input"
                type="file"
                accept=".xlsx,.xls"
                hidden
                onChange={(e) => processFile(e.target.files?.[0])}
              />
              <div className="upload-icon"><Upload size={28} /></div>
              <h2>Загрузите Excel-файл</h2>
              <p>Перетащите файл сюда или нажмите, чтобы выбрать</p>
              <span className="hint">Формат: .xlsx / .xls</span>
            </section>
          )}

          {file && (
            <>
              <section className="file-card">
                <div className="file-info">
                  <div className="file-icon"><FileSpreadsheet size={22} /></div>
                  <div>
                    <strong>{file.name}</strong>
                    <span>{(file.size / 1024).toFixed(1)} KB</span>
                  </div>
                </div>
                <button className="icon-btn" onClick={reset} title="Удалить файл">
                  <X size={19} />
                </button>
              </section>

              {loading && <div className="processing">Обрабатываем файл…</div>}

              {!loading && (results.length > 0 || errors.length > 0) && (
                <>
                  <section className="stats">
                    <div><b>{stats.total}</b><span>строк</span></div>
                    <div><b>{stats.success}</b><span>расшифровано</span></div>
                    <div className={stats.failed ? "bad" : ""}><b>{stats.failed}</b><span>ошибок</span></div>
                  </section>

                  {errors.length > 0 && (
                    <div className="warning">
                      <AlertCircle size={20} />
                      <div>
                        <strong>Есть строки, которые не удалось расшифровать</strong>
                        <p>Они будут вынесены на отдельный лист «Ошибки» в итоговом Excel.</p>
                      </div>
                    </div>
                  )}

                  <section className="preview-card">
                    <div className="section-title">
                      <div>
                        <h2>Предпросмотр</h2>
                        <p>Первые {Math.min(5, results.length)} успешно обработанных строк</p>
                      </div>
                      <CheckCircle2 className="success-icon" size={22} />
                    </div>

                    <div className="table-wrap">
                      <table>
                        <thead>
                          <tr>
                            <th>ФИО</th>
                            {OUT_ORDER.slice(0, 3).map(m => <th key={m}>{m}</th>)}
                            <th>…</th>
                          </tr>
                        </thead>
                        <tbody>
                          {results.slice(0, 5).map((r, i) => (
                            <tr key={i}>
                              <td>{r["ФИО"]}</td>
                              {OUT_ORDER.slice(0, 3).map(m => (
                                <td key={m}>
                                  <b>{r[m]}</b>
                                  <span className="level">{r["уровень_" + m]}</span>
                                </td>
                              ))}
                              <td>…</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </section>

                  <button
                    className="download-btn"
                    onClick={() => exportExcel(results, errors)}
                    disabled={!results.length}
                  >
                    <Download size={20} />
                    Скачать результаты Excel
                  </button>

                  <button className="secondary-btn" onClick={reset}>
                    Обработать другой файл
                  </button>
                </>
              )}
            </>
          )}
        </main>

        <footer>
          Данные обрабатываются прямо в браузере и не загружаются на сервер.
        </footer>
      </div>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
