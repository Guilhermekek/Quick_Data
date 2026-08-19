import React, { useEffect, useState } from "react";

function usePywebviewReady() {
  const [ready, setReady] = useState(!!window.pywebview);
  useEffect(() => {
    if (ready) return;
    const onReady = () => setReady(true);
    window.addEventListener("pywebviewready", onReady);
    return () => window.removeEventListener("pywebviewready", onReady);
  }, [ready]);
  return ready;
}

export default function ImportarFront() {
  const ready = usePywebviewReady();
  const [sourcePath, setSourcePath] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [selected, setSelected] = useState(new Set());
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState(null); // { kind: "ok"|"err", text }

  async function escolherArquivo() {
    setMessage(null);
    const path = await window.pywebview.api.pick_file();
    if (!path) return;

    setBusy(true);
    try {
      const json = await window.pywebview.api.list_sheets(path);
      const list = JSON.parse(json);
      setSourcePath(path);
      setSheets(list);
      setSelected(new Set(list.filter((s) => s.is_front).map((s) => s.name)));
    } catch (err) {
      setMessage({ kind: "err", text: "Não consegui ler esse arquivo: " + err });
    } finally {
      setBusy(false);
    }
  }

  function toggle(name) {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
  }

  async function importar() {
    if (selected.size === 0) return;
    setMessage(null);
    const dest = await window.pywebview.api.pick_dest_file();
    if (!dest) return;

    setBusy(true);
    try {
      const json = await window.pywebview.api.import_fronts(
        sourcePath,
        Array.from(selected),
        dest
      );
      const result = JSON.parse(json);
      setMessage({
        kind: "ok",
        text:
          `${result.created.length} aba(s) importada(s) para ${dest}: ` +
          result.created.join(", "),
      });
    } catch (err) {
      setMessage({ kind: "err", text: "Falha ao importar: " + err });
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="screen">
      <h1>Extração de Dados — Importar Front</h1>
      <p className="subtitle">
        Escolha um arquivo Quick Data (ou similar) e selecione quais abas de
        relatório ("Fronts") importar. Abas de infraestrutura (Base, tabelas
        mestre, config) já vêm desmarcadas.
      </p>

      <div className="card">
        <button className="btn btn-primary" onClick={escolherArquivo} disabled={!ready || busy}>
          {sourcePath ? "Trocar arquivo" : "Escolher arquivo"}
        </button>
        {sourcePath && (
          <span style={{ marginLeft: 12, color: "var(--fg-muted)", fontSize: 12 }}>
            {sourcePath}
          </span>
        )}
        {!ready && (
          <div className="toast toast-err" style={{ marginTop: 14 }}>
            Ponte com o Python ainda não carregou — normal por 1-2s ao abrir o app.
          </div>
        )}

        {sheets.length > 0 && (
          <>
            <table className="sheet-list">
              <thead>
                <tr>
                  <th></th>
                  <th>Aba</th>
                  <th>Tipo</th>
                  <th>Linhas × Colunas</th>
                </tr>
              </thead>
              <tbody>
                {sheets.map((s) => (
                  <tr key={s.name}>
                    <td>
                      <input
                        type="checkbox"
                        checked={selected.has(s.name)}
                        onChange={() => toggle(s.name)}
                      />
                    </td>
                    <td>{s.name}</td>
                    <td>
                      <span className={"badge " + (s.is_front ? "badge-front" : "badge-infra")}>
                        {s.is_front ? "Front" : "Infraestrutura"}
                      </span>
                    </td>
                    <td>
                      {s.rows} × {s.cols}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div style={{ marginTop: 16 }}>
              <button
                className="btn btn-primary"
                onClick={importar}
                disabled={busy || selected.size === 0}
              >
                Importar {selected.size} aba(s) selecionada(s)
              </button>
            </div>
          </>
        )}

        {message && (
          <div className={"toast " + (message.kind === "ok" ? "toast-ok" : "toast-err")}>
            {message.text}
          </div>
        )}
      </div>
    </div>
  );
}
