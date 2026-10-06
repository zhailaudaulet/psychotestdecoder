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
  BrainCircuit,
  Compass,
} from "lucide-react";
import {
  TEST_DEFINITIONS,
  decodeTestUrl,
  getLevel,
} from "./decoder.js";
import "./styles.css";

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

function downloadDecoderTemplate(testDefinition) {
  const data = [
    ["ФИО", "ссылка"],
    [
      "Иванов Иван Иванович",
      testDefinition.exampleUrl,
    ],
    [
      "Петров Петр Петрович",
      testDefinition.exampleUrl,
    ],
    [
      "Сидорова Анна Сергеевна",
      testDefinition.exampleUrl,
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
    `${testDefinition.label.toLowerCase()}_template.xlsx`
  );
}

/* =========================
   DECODER EXPORT
========================= */

function downloadDecoderResults(testDefinition, results, errors) {
  const headers = [
    "ФИО",
    ...testDefinition.metrics.flatMap((name) =>
      testDefinition.includeLevels ? [name, "уровень"] : [name]
    ),
  ];

  const rows = results.map((result) => {
    return [
      result["ФИО"],
      ...testDefinition.metrics.flatMap((name) =>
        testDefinition.includeLevels
          ? [result[name], result[`level_${name}`]]
          : [result[name]]
      ),
    ];
  });

  const worksheet = XLSX.utils.aoa_to_sheet([
    headers,
    ...rows,
  ]);

  worksheet["!cols"] = [
    { wch: 32 },
    ...headers.slice(1).map(() => ({ wch: 22 })),
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
    testDefinition.fileName
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

function createDecoderStates() {
  return Object.fromEntries(
    Object.keys(TEST_DEFINITIONS).map((testId) => [
      testId,
      { file: null, results: [], errors: [], loading: false },
    ])
  );
}

/* =========================
   APP
========================= */

function App() {
  const [tab, setTab] = useState("decoder");

  /* Decoder */

  const [activeTestId, setActiveTestId] = useState("ito");
  const [decoderStates, setDecoderStates] = useState(createDecoderStates);
  const [decoderDragActive, setDecoderDragActive] = useState(false);
  const activeTest = TEST_DEFINITIONS[activeTestId];
  const decoderState = decoderStates[activeTestId];

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

  function updateDecoderState(testId, changes) {
    setDecoderStates((current) => ({
      ...current,
      [testId]: {
        ...current[testId],
        ...changes,
      },
    }));
  }

  async function processDecoder(file, testId = activeTestId) {
    if (!file) return;

    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      alert(
        "Загрузите Excel-файл .xlsx или .xls"
      );
      return;
    }

    const testDefinition = TEST_DEFINITIONS[testId];

    updateDecoderState(testId, {
      file,
      loading: true,
      results: [],
      errors: [],
    });

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
          const decoded = decodeTestUrl(url, testId);

          const result = {
            ФИО: name,
          };

          for (const metric of testDefinition.metrics) {
            result[metric] =
              decoded[metric];

            if (testDefinition.includeLevels) {
              result[`level_${metric}`] =
                getLevel(decoded[metric]);
            }
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

      updateDecoderState(testId, { results, errors });
    } catch (error) {
      alert(error.message);
      updateDecoderState(testId, { file: null });
    } finally {
      updateDecoderState(testId, { loading: false });
    }
  }

  function resetDecoder(testId = activeTestId) {
    updateDecoderState(testId, {
      file: null,
      results: [],
      errors: [],
      loading: false,
    });
  }

  function handleDecoderDrag(event) {
    event.preventDefault();
    event.stopPropagation();
    event.dataTransfer.dropEffect = "copy";
    setDecoderDragActive(true);
  }

  function handleDecoderDragLeave(event) {
    event.preventDefault();
    event.stopPropagation();

    if (event.currentTarget.contains(event.relatedTarget)) return;

    setDecoderDragActive(false);
  }

  function handleDecoderDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    setDecoderDragActive(false);
    processDecoder(event.dataTransfer.files?.[0], activeTestId);
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
          <div className="header-identity">
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
          </div>

          <div className="orientation-visual" aria-hidden="true">
            <div className="orientation-orbit" />

            <div className="orientation-ring orientation-ring-primary">
              <Compass size={28} />
            </div>

            <div className="orientation-ring orientation-ring-secondary">
              <BrainCircuit size={23} />
            </div>
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

            <nav className="test-tabs" aria-label="Выбор психологического теста">
              {Object.entries(TEST_DEFINITIONS).map(([testId, definition]) => (
                <button
                  key={testId}
                  type="button"
                  className={activeTestId === testId ? "active" : ""}
                  aria-current={activeTestId === testId ? "page" : undefined}
                  onClick={() => setActiveTestId(testId)}
                >
                  <span>{definition.label}</span>
                  <small>{definition.metrics.length} шкал</small>
                </button>
              ))}
            </nav>

            <div className="page-title">

              <div>
                <h2>
                  {activeTest.title}
                </h2>

                <p>
                  {activeTest.description}
                </p>
              </div>

              <button
                type="button"
                className="outline-btn template-btn"
                onClick={() => downloadDecoderTemplate(activeTest)}
              >
                <Download size={17} />
                Шаблон {activeTest.label}
              </button>

            </div>

            {/* TEMPLATE PREVIEW */}

            <section className="demo-card">

              <div className="demo-header">
                <div>
                  <strong>
                    Шаблон для теста «{activeTest.label}»
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

              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>ФИО</th>
                      <th>ссылка</th>
                    </tr>
                  </thead>

                  <tbody>
                    <tr>
                      <td>Иванов Иван Иванович</td>
                      <td>{activeTest.exampleUrl}</td>
                    </tr>
                    <tr>
                      <td>Петров Петр Петрович</td>
                      <td>{activeTest.exampleUrl}</td>
                    </tr>
                    <tr>
                      <td>Сидорова Анна Сергеевна</td>
                      <td>{activeTest.exampleUrl}</td>
                    </tr>
                  </tbody>
                </table>
              </div>

            </section>

            {/* UPLOAD */}

            {!decoderState.file && (
              <label
                className={`dropzone ${decoderDragActive ? "drag-active" : ""}`}
                onDragEnter={handleDecoderDrag}
                onDragOver={handleDecoderDrag}
                onDragLeave={handleDecoderDragLeave}
                onDrop={handleDecoderDrop}
              >

                <input
                  hidden
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(e) => {
                    setDecoderDragActive(false);
                    processDecoder(
                      e.target.files[0],
                      activeTestId
                    );
                  }}
                />

                <div className="upload-icon">
                  <Upload size={28} />
                </div>

                <h2>
                  Загрузите файл «{activeTest.label}»
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

            {decoderState.file && (
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
                        {decoderState.file.name}
                      </strong>

                      <span>
                        Входной файл · {activeTest.label}
                      </span>
                    </div>

                  </div>

                  <button
                    type="button"
                    className="icon-btn"
                    aria-label={`Удалить файл ${activeTest.label}`}
                    onClick={() => resetDecoder(activeTestId)}
                  >
                    <X size={18} />
                  </button>

                </div>

                {decoderState.loading && (
                  <div className="processing" role="status">
                    Расшифровываем ссылки…
                  </div>
                )}

                {!decoderState.loading &&
                  (decoderState.results.length > 0 ||
                    decoderState.errors.length > 0) && (
                    <>

                      <div className="stats">

                        <div>
                          <b>
                            {decoderState.results.length}
                          </b>

                          <span>
                            расшифровано
                          </span>
                        </div>

                        <div>
                          <b>
                            {decoderState.errors.length}
                          </b>

                          <span>
                            ошибок
                          </span>
                        </div>

                        <div>
                          <b>
                            {decoderState.results.length +
                              decoderState.errors.length}
                          </b>

                          <span>
                            всего строк
                          </span>
                        </div>

                      </div>

                      {decoderState.errors.length >
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

                      {decoderState.results.length > 0 && (
                        <ResultsPreview
                          testDefinition={activeTest}
                          results={decoderState.results}
                        />
                      )}

                      <button
                        type="button"
                        className="download-btn"
                        disabled={
                          !decoderState.results.length
                        }
                        onClick={() =>
                          downloadDecoderResults(
                            activeTest,
                            decoderState.results,
                            decoderState.errors
                          )
                        }
                      >
                        <Download size={19} />
                        Скачать результаты {activeTest.label}
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
   RESULTS PREVIEW
========================= */

function ResultsPreview({ testDefinition, results }) {
  const previewRows = results.slice(0, 5);

  return (
    <section className="results-preview">
      <div className="preview-header">
        <div>
          <h3>Предпросмотр результатов</h3>
          <p>
            Показано {previewRows.length} из {results.length} строк
          </p>
        </div>

        <FileSpreadsheet size={20} />
      </div>

      <div className="table-wrap results-table-wrap">
        <table className="results-table">
          <thead>
            <tr>
              <th>ФИО</th>
              {testDefinition.metrics.map((metric) => (
                <th key={metric}>{metric}</th>
              ))}
            </tr>
          </thead>

          <tbody>
            {previewRows.map((result, index) => (
              <tr key={`${result["ФИО"]}-${index}`}>
                <td className="result-name">
                  {result["ФИО"] || "Без имени"}
                </td>

                {testDefinition.metrics.map((metric) => (
                  <td key={metric}>
                    <strong className="result-score">
                      {result[metric]}
                    </strong>

                    {testDefinition.includeLevels && (
                      <span className="result-level">
                        {result[`level_${metric}`]}
                      </span>
                    )}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
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
  const [isDragging, setIsDragging] = useState(false);

  function selectFile(selectedFile) {
    if (!selectedFile) return;

    if (!/\.(xlsx|xls)$/i.test(selectedFile.name)) {
      alert("Загрузите Excel-файл .xlsx или .xls");
      return;
    }

    setFile(selectedFile);
  }

  function handleDrag(event) {
    event.preventDefault();
    event.stopPropagation();
    event.dataTransfer.dropEffect = "copy";
    setIsDragging(true);
  }

  function handleDragLeave(event) {
    event.preventDefault();
    event.stopPropagation();

    if (event.currentTarget.contains(event.relatedTarget)) return;

    setIsDragging(false);
  }

  function handleDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(false);
    selectFile(event.dataTransfer.files?.[0]);
  }

  return (
    <section
      className={`upload-card ${isDragging ? "drag-active" : ""}`}
      onDragEnter={handleDrag}
      onDragOver={handleDrag}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >

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
              type="button"
              aria-label={`Удалить файл ${file.name}`}
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
              onChange={(e) => selectFile(e.target.files[0])}
            />

            <Upload size={17} />

            Выбрать или перетащить Excel

          </label>
        )}

      </div>

    </section>
  );
}

createRoot(
  document.getElementById("root")
).render(<App />);
