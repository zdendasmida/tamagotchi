import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const FILTER_OPERATORS = [
  { value: "contains", label: "Obsahuje" },
  { value: "equals", label: "Je rovno" },
  { value: "startsWith", label: "Začíná na" },
  { value: "greaterThan", label: "Větší než" },
  { value: "lessThan", label: "Menší než" }
];

function normalize(value) {
  if (value === undefined || value === null) {
    return "";
  }
  if (typeof value === "number") {
    return value;
  }
  return String(value);
}

function matchesFilter(row, column, operator, rawFilterValue) {
  if (!column || rawFilterValue.trim() === "") {
    return true;
  }

  const cellValue = row[column];
  const normalizedCell = normalize(cellValue);
  const normalizedFilter = normalize(rawFilterValue);

  switch (operator) {
    case "equals":
      if (typeof normalizedCell === "number" && !Number.isNaN(Number(normalizedFilter))) {
        return normalizedCell === Number(normalizedFilter);
      }
      return String(normalizedCell).toLowerCase() === String(normalizedFilter).toLowerCase();
    case "startsWith":
      return String(normalizedCell).toLowerCase().startsWith(String(normalizedFilter).toLowerCase());
    case "greaterThan": {
      const cellNumber = Number(normalizedCell);
      const filterNumber = Number(normalizedFilter);
      if (Number.isNaN(cellNumber) || Number.isNaN(filterNumber)) {
        return false;
      }
      return cellNumber > filterNumber;
    }
    case "lessThan": {
      const cellNumber = Number(normalizedCell);
      const filterNumber = Number(normalizedFilter);
      if (Number.isNaN(cellNumber) || Number.isNaN(filterNumber)) {
        return false;
      }
      return cellNumber < filterNumber;
    }
    case "contains":
    default:
      return String(normalizedCell).toLowerCase().includes(String(normalizedFilter).toLowerCase());
  }
}

export default function App() {
  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [columns, setColumns] = useState([]);
  const [filterColumn, setFilterColumn] = useState("");
  const [filterOperator, setFilterOperator] = useState("contains");
  const [filterValue, setFilterValue] = useState("");
  const [error, setError] = useState("");

  const activeSheet = useMemo(
    () => sheets.find((sheet) => sheet.name === selectedSheet) ?? null,
    [sheets, selectedSheet]
  );

  const rows = activeSheet?.rows ?? [];
  const totalRows = rows.length;

  const filteredRows = useMemo(() => {
    if (!rows.length) {
      return [];
    }

    return rows.filter((row) => matchesFilter(row, filterColumn, filterOperator, filterValue));
  }, [rows, filterColumn, filterOperator, filterValue]);

  const handleFileUpload = (event) => {
    const file = event.target.files?.[0];
    setError("");

    if (!file) {
      return;
    }

    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const buffer = e.target?.result;
        const data = new Uint8Array(buffer);
        const workbook = XLSX.read(data, { type: "array" });

        const nextSheets = workbook.SheetNames.map((name) => {
          const worksheet = workbook.Sheets[name];
          const rows = XLSX.utils.sheet_to_json(worksheet, {
            defval: "",
            raw: true
          });
          return { name, rows };
        });

        setSheets(nextSheets);
        const firstSheet = nextSheets[0];
        const firstColumns = firstSheet ? Object.keys(firstSheet.rows[0] ?? {}) : [];
        setColumns(firstColumns);
        setSelectedSheet(firstSheet?.name ?? "");
        setFilterColumn(firstColumns[0] ?? "");
        setFilterOperator("contains");
        setFilterValue("");
      } catch (uploadError) {
        console.error(uploadError);
        setError("Nepodařilo se načíst soubor. Zkontrolujte prosím, že jde o soubor XLSX/XLS/CSV.");
      }
    };

    reader.onerror = () => {
      setError("Během čtení souboru došlo k chybě.");
    };

    reader.readAsArrayBuffer(file);
  };

  const handleSheetChange = (event) => {
    const sheetName = event.target.value;
    setSelectedSheet(sheetName);
    const newSheet = sheets.find((sheet) => sheet.name === sheetName);
    const newColumns = newSheet ? Object.keys(newSheet.rows[0] ?? {}) : [];
    setColumns(newColumns);
    setFilterColumn(newColumns[0] ?? "");
    setFilterValue("");
    setFilterOperator("contains");
  };

  const resetFilter = () => {
    setFilterValue("");
    setFilterOperator("contains");
  };

  return (
    <main className="app">
      <header>
        <h1>Filtr pro Google Sheet</h1>
        <p>
          Nahrajte export Google Sheetu (XLSX, XLS nebo CSV) a poté si vyberte list a filtr, který se má
          použít na data.
        </p>
      </header>

      <section className="panel">
        <div className="input-group">
          <label htmlFor="sheet-file">Soubor</label>
          <input id="sheet-file" type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} />
        </div>

        {Boolean(sheets.length) && (
          <>
            <div className="input-group">
              <label htmlFor="sheet-select">List</label>
              <select id="sheet-select" value={selectedSheet} onChange={handleSheetChange}>
                {sheets.map((sheet) => (
                  <option key={sheet.name} value={sheet.name}>
                    {sheet.name}
                  </option>
                ))}
              </select>
            </div>

            <div className="filters">
              <div className="input-group">
                <label htmlFor="column-select">Sloupec</label>
                <select
                  id="column-select"
                  value={filterColumn}
                  onChange={(event) => setFilterColumn(event.target.value)}
                >
                  {columns.length ? (
                    columns.map((column) => (
                      <option key={column} value={column}>
                        {column || "(prázdný)"}
                      </option>
                    ))
                  ) : (
                    <option value="" disabled>
                      Žádné sloupce
                    </option>
                  )}
                </select>
              </div>

              <div className="input-group">
                <label htmlFor="operator-select">Operátor</label>
                <select
                  id="operator-select"
                  value={filterOperator}
                  onChange={(event) => setFilterOperator(event.target.value)}
                >
                  {FILTER_OPERATORS.map((operator) => (
                    <option key={operator.value} value={operator.value}>
                      {operator.label}
                    </option>
                  ))}
                </select>
              </div>

              <div className="input-group">
                <label htmlFor="filter-value">Hodnota</label>
                <input
                  id="filter-value"
                  type="text"
                  value={filterValue}
                  onChange={(event) => setFilterValue(event.target.value)}
                  placeholder="Zadejte hodnotu filtru"
                />
              </div>

              <button className="secondary" type="button" onClick={resetFilter}>
                Vymazat filtr
              </button>
            </div>
          </>
        )}

        {error && <div className="error">{error}</div>}
      </section>

      <section className="table-wrapper">
        {Boolean(totalRows) && (
          <div className="table-summary">
            Zobrazeno {filteredRows.length} z {totalRows} řádků
          </div>
        )}

        {totalRows === 0 ? (
          sheets.length ? (
            <div className="empty">Vybraný list neobsahuje žádná data.</div>
          ) : (
            <div className="empty">Nejprve nahrajte soubor Google Sheetu.</div>
          )
        ) : filteredRows.length ? (
          <div className="table-scroll">
            <table>
              <thead>
                <tr>
                  {columns.map((column) => (
                    <th key={column}>{column || "(prázdný)"}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredRows.map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {columns.map((column) => (
                      <td key={column}>{normalize(row[column])}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <div className="empty">Žádné řádky neodpovídají zadanému filtru.</div>
        )}
      </section>
    </main>
  );
}
