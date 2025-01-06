import { useState } from "react";
import "./App.css";
import { Button } from "./components/ui/button";
import { Textarea } from "./components/ui/textarea";
import { getCommentedModule } from "./vba/vbeTextConverters";

function App() {
  const [inputText, setInputText] = useState("");
  const [outputText, setoutputText] = useState("");

  function handleClearAll() {
    setInputText("");
    setoutputText("");
  }

  function handleConvert() {
    setoutputText(getCommentedModule(inputText));
  }

  function handleCopy() {
    navigator.clipboard.writeText(outputText);
    alert("Copied output!");
  }

  return (
    <>
      <div className="flex flex-col bg-slate-300 text-slate-900 min-h-screen p-4">
        <header className="mx-auto">
          <h1 className="font-bold text-4xl">VBA Converter</h1>
        </header>
        <main className="flex grow justify-center gap-4">
          <div className="flex flex-col grow gap-4">
            <h2 className="text-center font-bold text-2xl">Input</h2>
            <div className="flex flex-col grow gap-4">
              <Textarea
                className="grow"
                onChange={(e) => setInputText(e.target.value)}
                value={inputText}
                spellCheck="false"
              />
              <Button onClick={handleConvert}>Convert</Button>
              <Button onClick={handleClearAll}>Clear All</Button>
            </div>
          </div>
          <div className="flex flex-col grow gap-4">
            <h2 className="text-center font-bold text-2xl">Output</h2>
            <div className="flex flex-col grow gap-4">
              <Textarea className="grow" value={outputText} readOnly />
              <Button onClick={handleCopy}>Copy Output</Button>
            </div>
          </div>
        </main>
      </div>
    </>
  );
}

export default App;
