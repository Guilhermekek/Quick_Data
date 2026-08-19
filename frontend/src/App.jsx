import React, { useState } from "react";
import PainelPrincipal from "./screens/PainelPrincipal.jsx";
import ImportarFront from "./screens/ImportarFront.jsx";

const NAV = [
  { id: "painel", label: "Painel Principal" },
  { id: "importar", label: "Extração / Importar Front" },
];

export default function App() {
  const [screen, setScreen] = useState("importar");

  return (
    <div className="app">
      <aside className="sidebar">
        <div className="logo-mark">
          <span className="tri" />
          <span>Quick Data</span>
        </div>
        <nav>
          {NAV.map((item) => (
            <div
              key={item.id}
              className={"nav-item" + (screen === item.id ? " active" : "")}
              onClick={() => setScreen(item.id)}
            >
              {item.label}
            </div>
          ))}
        </nav>
      </aside>

      <main className="main">
        {screen === "painel" && <PainelPrincipal />}
        {screen === "importar" && <ImportarFront />}
      </main>

      <footer className="footer">
        <span>Planning &amp; Control · Quick Data v0.1.0-dev</span>
        <span>Guilherme Kek</span>
      </footer>
    </div>
  );
}
