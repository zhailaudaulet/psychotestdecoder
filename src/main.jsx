import React, { useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Upload,
  Download,
  FileSpreadsheet,
  Merge,
  AlertCircle,
  X,
} from "lucide-react";
import "./styles.css";

/* =========================
   DECODER LOGIC
========================= */

const ALPHABET =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

const MODULUS = 10 ** 10;
const OFFSET = 3089424140;

const ORDER_LOW_TO_HIGH = [
  "Шкала искренности",
  "Лабильность",
  "Тревожность",
  "Сензитивность",
  "Интроверсия",
  "Ригидность",
  "Агрессивность",
  "Спонтанность",
  "Экстраверсия",
  "Шкала аггравации",
];

const OUTPUT_ORDER = [
  "Экстраверсия",
  "Спонтанность",
  "Агрессивность",
  "Ригидность",
  "Интроверсия",
  "Сензитивность",
  "Тревожность",
  "Лабильность",
  "Шкала искренности",
  "Шкала аггравации",
];

function charValue(char) {
  const value = ALPHABET.indexOf(char);

  if (value === -1) {
    throw new Error(`Недопустимый символ: ${char}`);
  }

  return value >= 52 ? value - 64 : value;
}

function decode(payload) {
  let n = 0;

  for (const char of payload) {
    n = n * 64 + charValue(char);
  }

  n = (n + OFFSET) % MODULUS;

  const digits = String(n)
    .padStart(10, "0")
    .split("")
    .reverse()
    .map(Number);

  return Object.fromEntries(
    ORDER_LOW_TO_HIGH.map((name, index) => [name, digits[index]])
  );
}

function getLevel(score) {
  if (score <= 1) return "гипоэмотивность";
  if (score <= 4) return "норма";
  if (score <= 7) return "акцентуация";
  return "избыточность";
}

function decodeUrl(url) {
  const match = String(url).match(/[?&]v=([A-Za-z0-9_-]+)/);

  if (!match) {
    throw new Error("Не найден параметр v");
  }

  const payload = match[1];

  if (!payload.startsWith("ito") || payload.length !== 10) {
    throw new Error("Неожиданный формат ссылки");
  }

  const code = payload.slice(3);

  if (code[0] !== "A") {
    throw new Error("Неподдерживаемая версия теста");
  }

  return decode(code.slice(1));
}

/* =========================
   EXCEL HELPERS
========================= */

async function readExcel(file) {
  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array",
  });

  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  return XLSX.utils.sheet_to_json(sheet, {
    defval: "",
  });
}

function findColumn(rows, name) {
  if (!rows.length) return null;

  return Object.keys(rows[0]).find(
    (key) => String(key).trim().toLowerCase() === name.toLowerCase()
  );
}

function normalizeName(value) {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleLowerCase("ru-RU");
}

/* =========================
   TEMPLATE
========================= */

function downloadDecoderTemplate() {
  const data = [
    ["ФИО", "ссылка"],
    [
      "Иванов Иван Иванович",
      "https://psytests.org/result?v=itoAItqccE",
    ],
    [
      "Петров Петр Петрович",
      "https://psytests.org/result?v=itoAItqccE",
    ],
    [
      "Сидорова Анна Сергеевна",
      "https://psytests.org/result?v=itoAItqccE",
    ],
  ];

  const workbook = XLSX.utils.book_new();

  const sheet = XLSX.utils.aoa_to_sheet(data);

  sheet["!cols"] = [
    { wch: 32 },
    { wch: 65 },
  ];

  XLSX.utils.book_append_sheet(
    workbook,
    sheet,
    "Decoder"
  );

  XLSX.writeFile(
    workbook,
    "decoder_template.xlsx"
  );
}

/* =========================
   DECODER EXPORT
========================= */

function downloadDecoderResults(results, errors) {
  const headers = [
    "ФИО",
    ...OUTPUT_ORDER.flatMap((name) => [
      name,
      "уровень",
    ]),
  ];

  const rows = results.map((result) => {
    return [
      result["ФИО"],
      ...OUTPUT_ORDER.flatMap((name) => [
        result[name],
        result[`level_${name}`],
      ]),
    ];
  });

  const worksheet = XLSX.utils.aoa_to_sheet([
    headers,
    ...rows,
  ]);

  worksheet["!cols"] = [
    { wch: 32 },
    ...Array(20).fill({ wch: 19 }),
  ];

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Результаты"
  );

  if (errors.length > 0) {
    const errorRows = [
      ["ФИО", "Ссылка", "Ошибка"],
      ...errors.map((error) => [
        error.name,
        error.url,
        error.error,
      ]),
    ];

    const errorSheet =
      XLSX.utils.aoa_to_sheet(errorRows);

    errorSheet["!cols"] = [
      { wch: 32 },
      { wch: 65 },
      { wch: 45 },
    ];

    XLSX.utils.book_append_sheet(
      workbook,
      errorSheet,
      "Ошибки"
    );
  }

  XLSX.writeFile(
    workbook,
    "psytest_results.xlsx"
  );
}

/* =========================
   MERGER EXPORT
========================= */

function downloadMergedResults(rows) {
  if (!rows.length) return;

  const worksheet =
    XLSX.utils.json_to_sheet(rows);

  worksheet["!cols"] = Object.keys(rows[0]).map(
    (key) => ({
      wch: Math.max(
        15,
        Math.min(40, key.length + 5)
      ),
    })
  );

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Результаты"
  );

  XLSX.writeFile(
    workbook,
    "psytest_merged.xlsx"
  );
}

/* =========================
   APP
========================= */

function App() {
  const [tab, setTab] = useState("decoder");

  /* Decoder */

  const [decoderFile, setDecoderFile] =
    useState(null);

  const [decoderResults, setDecoderResults] =
    useState([]);

  const [decoderErrors, setDecoderErrors] =
    useState([]);

  const [decoderLoading, setDecoderLoading] =
    useState(false);

  /* Merger */

  const [resultFile, setResultFile] =
    useState(null);

  const [attendanceFile, setAttendanceFile] =
    useState(null);

  const [mergedResults, setMergedResults] =
    useState([]);

  const [mergeErrors, setMergeErrors] =
    useState([]);

  const [mergeLoading, setMergeLoading] =
    useState(false);

  /* =========================
     DECODER
  ========================= */

  async function processDecoder(file) {
    if (!file) return;

    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      alert(
        "Загрузите Excel-файл .xlsx или .xls"
      );
      return;
    }

    setDecoderFile(file);
    setDecoderLoading(true);
    setDecoderResults([]);
    setDecoderErrors([]);

    try {
      const data = await readExcel(file);

      if (!data.length) {
        throw new Error(
          "Excel-файл пустой"
        );
      }

      const nameColumn =
        findColumn(data, "ФИО");

      const urlColumn =
        findColumn(data, "ссылка");

      if (!nameColumn || !urlColumn) {
        throw new Error(
          'В файле должны быть столбцы "ФИО" и "ссылка"'
        );
      }

      const results = [];
      const errors = [];

      for (const row of data) {
        const name = String(
          row[nameColumn] ?? ""
        ).trim();

        const url = String(
          row[urlColumn] ?? ""
        ).trim();

        if (!name && !url) continue;

        try {
          const decoded = decodeUrl(url);

          const result = {
            ФИО: name,
          };

          for (const metric of OUTPUT_ORDER) {
            result[metric] =
              decoded[metric];

            result[`level_${metric}`] =
              getLevel(decoded[metric]);
          }

          results.push(result);
        } catch (error) {
          errors.push({
            name,
            url,
            error: error.message,
          });
        }
      }

      setDecoderResults(results);
      setDecoderErrors(errors);
    } catch (error) {
      alert(error.message);
      setDecoderFile(null);
    } finally {
      setDecoderLoading(false);
    }
  }

  /* =========================
     MERGER
  ========================= */

  async function processMerger() {
    if (!resultFile || !attendanceFile) {
      alert(
        "Загрузите оба файла"
      );
      return;
    }

    setMergeLoading(true);
    setMergedResults([]);
    setMergeErrors([]);

    try {
      const [
        decoderData,
        attendanceData,
      ] = await Promise.all([
        readExcel(resultFile),
        readExcel(attendanceFile),
      ]);

      if (!decoderData.length) {
        throw new Error(
          "Файл Decoder пустой"
        );
      }

      if (!attendanceData.length) {
        throw new Error(
          "Файл посещения пустой"
        );
      }

      const decoderNameColumn =
        findColumn(decoderData, "ФИО");

      const attendanceNameColumn =
        findColumn(
          attendanceData,
          "ФИО"
        );

      const attendanceColumn =
        findColumn(
          attendanceData,
          "Посещение"
        );

      if (!decoderNameColumn) {
        throw new Error(
          'В первом файле нет столбца "ФИО"'
        );
      }

      if (
        !attendanceNameColumn ||
        !attendanceColumn
      ) {
        throw new Error(
          'Во втором файле должны быть столбцы "ФИО" и "Посещение"'
        );
      }

      /*
       * Создаём индекс:
       *
       * "иванов иван иванович"
       * →
       * "Присутствовал"
       */

      const attendanceMap = new Map();

      for (const row of attendanceData) {
        const name = normalizeName(
          row[attendanceNameColumn]
        );

        if (!name) continue;

        attendanceMap.set(
          name,
          row[attendanceColumn]
        );
      }

      const matched = [];
      const unmatched = [];

      /*
       * Идём именно по Decoder-файлу.
       * Поэтому все его результаты сохраняют
       * свою структуру.
       */

      for (const row of decoderData) {
        const name = String(
          row[decoderNameColumn] ?? ""
        );

        const normalized =
          normalizeName(name);

        if (
          normalized &&
          attendanceMap.has(normalized)
        ) {
          matched.push({
            ...row,
            Посещение:
              attendanceMap.get(normalized),
          });
        } else {
          unmatched.push({
            name,
            reason:
              "Не найдено совпадение во втором файле",
          });
        }
      }

      setMergedResults(matched);
      setMergeErrors(unmatched);
    } catch (error) {
      alert(error.message);
    } finally {
      setMergeLoading(false);
    }
  }

  function resetMerger() {
    setResultFile(null);
    setAttendanceFile(null);
    setMergedResults([]);
    setMergeErrors([]);
  }

  return (
    <div className="page">
      <div className="container">

        {/* HEADER */}

        <header className="header">
          <div className="brand-mark">
            <FileSpreadsheet size={22} />
          </div>

          <div>
            <h1>
              PsychTest Tools
            </h1>

            <p>
              Расшифровка и объединение результатов
            </p>
          </div>
        </header>

        {/* TABS */}

        <nav className="tabs">

          <button
            className={
              tab === "decoder"
                ? "active"
                : ""
            }
            onClick={() =>
              setTab("decoder")
            }
          >
            Decoder
          </button>

          <button
            className={
              tab === "merger"
                ? "active"
                : ""
            }
            onClick={() =>
              setTab("merger")
            }
          >
            Zoom Merger
          </button>

        </nav>

        {/* =========================
            DECODER TAB
        ========================= */}

        {tab === "decoder" && (
          <main>

            <div className="page-title">

              <div>
                <h2>
                  Decoder
                </h2>

                <p>
                  Расшифровка ссылок из Excel
                </p>
              </div>

              <button
                className="outline-btn"
                onClick={
                  downloadDecoderTemplate
                }
              >
                <Download size={17} />
                Скачать шаблон
              </button>

            </div>

            {/* TEMPLATE PREVIEW */}

            <section className="demo-card">

              <div className="demo-header">
                <div>
                  <strong>
                    Шаблон входного файла
                  </strong>

                  <span>
                    Используйте столбцы
                    «ФИО» и «ссылка»
                  </span>
                </div>

                <FileSpreadsheet
                  size={20}
                />
              </div>

              <table>
                <thead>
                  <tr>
                    <th>ФИО</th>
                    <th>ссылка</th>
                  </tr>
                </thead>

                <tbody>

                  <tr>
                    <td>
                      Иванов Иван Иванович
                    </td>

                    <td>
                      https://psytests.org/result?v=itoAItqccE
                    </td>
                  </tr>

                  <tr>
                    <td>
                      Петров Петр Петрович
                    </td>

                    <td>
                      https://psytests.org/result?v=itoAItqccE
                    </td>
                  </tr>

                  <tr>
                    <td>
                      Сидорова Анна Сергеевна
                    </td>

                    <td>
                      https://psytests.org/result?v=itoAItqccE
                    </td>
                  </tr>

                </tbody>
              </table>

            </section>

            {/* UPLOAD */}

            {!decoderFile && (
              <label className="dropzone">

                <input
                  hidden
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(e) =>
                    processDecoder(
                      e.target.files[0]
                    )
                  }
                />

                <div className="upload-icon">
                  <Upload size={28} />
                </div>

                <h2>
                  Загрузите Excel-файл
                </h2>

                <p>
                  Нажмите или перетащите файл
                </p>

                <span>
                  .xlsx / .xls
                </span>

              </label>
            )}

            {/* FILE */}

            {decoderFile && (
              <>
                <div className="file-card">

                  <div className="file-info">

                    <div className="file-icon">
                      <FileSpreadsheet
                        size={22}
                      />
                    </div>

                    <div>
                      <strong>
                        {decoderFile.name}
                      </strong>

                      <span>
                        Входной файл
                      </span>
                    </div>

                  </div>

                  <button
                    className="icon-btn"
                    onClick={() => {
                      setDecoderFile(null);
                      setDecoderResults([]);
                      setDecoderErrors([]);
                    }}
                  >
                    <X size={18} />
                  </button>

                </div>

                {decoderLoading && (
                  <div className="processing">
                    Расшифровываем ссылки…
                  </div>
                )}

                {!decoderLoading &&
                  (decoderResults.length > 0 ||
                    decoderErrors.length > 0) && (
                    <>

                      <div className="stats">

                        <div>
                          <b>
                            {decoderResults.length}
                          </b>

                          <span>
                            расшифровано
                          </span>
                        </div>

                        <div>
                          <b>
                            {decoderErrors.length}
                          </b>

                          <span>
                            ошибок
                          </span>
                        </div>

                        <div>
                          <b>
                            {decoderResults.length +
                              decoderErrors.length}
                          </b>

                          <span>
                            всего строк
                          </span>
                        </div>

                      </div>

                      {decoderErrors.length >
                        0 && (
                        <div className="warning">

                          <AlertCircle
                            size={20}
                          />

                          <div>
                            <strong>
                              Есть ошибки
                            </strong>

                            <p>
                              Ошибочные строки
                              будут помещены
                              на отдельный лист
                              «Ошибки».
                            </p>
                          </div>

                        </div>
                      )}

                      <button
                        className="download-btn"
                        disabled={
                          !decoderResults.length
                        }
                        onClick={() =>
                          downloadDecoderResults(
                            decoderResults,
                            decoderErrors
                          )
                        }
                      >
                        <Download size={19} />
                        Скачать результаты Excel
                      </button>

                    </>
                  )}
              </>
            )}

          </main>
        )}

        {/* =========================
            ZOOM MERGER
        ========================= */}

        {tab === "merger" && (
          <main>

            <div className="page-title">

              <div>
                <h2>
                  Zoom Merger
                </h2>

                <p>
                  Объединение результатов
                  Decoder с посещением
                </p>
              </div>

            </div>

            {/* TWO FILES */}

            <div className="upload-grid">

              <UploadCard
                number="01"
                title="Результаты Decoder"
                description="Excel, скачанный из вкладки Decoder"
                file={resultFile}
                setFile={setResultFile}
                id="decoder-result"
              />

              <UploadCard
                number="02"
                title="Посещение"
                description='Excel со столбцами «ФИО» и «Посещение»'
                file={attendanceFile}
                setFile={setAttendanceFile}
                id="attendance"
              />

            </div>

            {/* ATTENDANCE DEMO */}

            <section className="demo-card">

              <div className="demo-header">
                <div>
                  <strong>
                    Формат файла посещения
                  </strong>

                  <span>
                    Сопоставление выполняется
                    по ФИО
                  </span>
                </div>
              </div>

              <table>

                <thead>
                  <tr>
                    <th>ФИО</th>
                    <th>Посещение</th>
                  </tr>
                </thead>

                <tbody>

                  <tr>
                    <td>
                      Иванов Иван Иванович
                    </td>

                    <td>
                      Присутствовал
                    </td>
                  </tr>

                  <tr>
                    <td>
                      Петров Петр Петрович
                    </td>

                    <td>
                      Отсутствовал
                    </td>
                  </tr>

                  <tr>
                    <td>
                      Сидорова Анна Сергеевна
                    </td>

                    <td>
                      Присутствовал
                    </td>
                  </tr>

                </tbody>

              </table>

            </section>

            {/* MERGE */}

            <button
              className="download-btn"
              disabled={
                !resultFile ||
                !attendanceFile ||
                mergeLoading
              }
              onClick={processMerger}
            >
              <Merge size={19} />

              {mergeLoading
                ? "Объединяем…"
                : "Объединить файлы"}
            </button>

            {/* RESULT */}

            {!mergeLoading &&
              (mergedResults.length > 0 ||
                mergeErrors.length > 0) && (
                <>

                  <div className="stats">

                    <div>
                      <b>
                        {mergedResults.length}
                      </b>

                      <span>
                        совпадений
                      </span>
                    </div>

                    <div>
                      <b>
                        {mergeErrors.length}
                      </b>

                      <span>
                        не найдено
                      </span>
                    </div>

                    <div>
                      <b>
                        {mergedResults.length +
                          mergeErrors.length}
                      </b>

                      <span>
                        строк
                      </span>
                    </div>

                  </div>

                  {/* UNMATCHED */}

                  {mergeErrors.length >
                    0 && (
                    <section className="errors-card">

                      <div className="section-title">

                        <div>
                          <h2>
                            Не найденные строки
                          </h2>

                          <p>
                            Эти ФИО отсутствуют
                            во втором файле
                          </p>
                        </div>

                        <AlertCircle
                          size={21}
                        />

                      </div>

                      <div className="table-wrap">

                        <table>

                          <thead>
                            <tr>
                              <th>ФИО</th>
                              <th>
                                Причина
                              </th>
                            </tr>
                          </thead>

                          <tbody>

                            {mergeErrors.map(
                              (error, index) => (
                                <tr
                                  key={index}
                                >
                                  <td>
                                    {error.name ||
                                      "—"}
                                  </td>

                                  <td>
                                    {error.reason}
                                  </td>
                                </tr>
                              )
                            )}

                          </tbody>

                        </table>

                      </div>

                    </section>
                  )}

                  {/* DOWNLOAD */}

                  <button
                    className="download-btn"
                    disabled={
                      !mergedResults.length
                    }
                    onClick={() =>
                      downloadMergedResults(
                        mergedResults
                      )
                    }
                  >
                    <Download size={19} />
                    Скачать объединённый Excel
                  </button>

                </>
              )}

          </main>
        )}

        <footer>
          Все данные обрабатываются локально
          в браузере и не отправляются
          на сервер.
        </footer>

      </div>
    </div>
  );
}

/* =========================
   UPLOAD CARD
========================= */

function UploadCard({
  number,
  title,
  description,
  file,
  setFile,
  id,
}) {
  return (
    <section className="upload-card">

      <div className="upload-number">
        {number}
      </div>

      <div className="upload-content">

        <h3>
          {title}
        </h3>

        <p>
          {description}
        </p>

        {file ? (
          <div className="selected-file">

            <FileSpreadsheet size={18} />

            <span>
              {file.name}
            </span>

            <button
              onClick={() =>
                setFile(null)
              }
            >
              <X size={16} />
            </button>

          </div>
        ) : (
          <label className="mini-upload">

            <input
              hidden
              type="file"
              accept=".xlsx,.xls"
              id={id}
              onChange={(e) =>
                setFile(
                  e.target.files[0]
                )
              }
            />

            <Upload size={17} />

            Выбрать Excel

          </label>
        )}

      </div>

    </section>
  );
}

createRoot(
  document.getElementById("root")
).render(<App />);