import React, { useEffect, useState } from "react";
import "./App.css";
import * as example from "./example";
import { DataModel } from "mr-excel";
import SharedExample from "./SharedExample";
function App() {
  const [examples, setExamples] = useState({});
  useEffect(() => {
    import("shared-example")
      .then((examples) => {
        setExamples(examples);
      })
      .catch((e) => {
        console.log(e);
      });
  }, []);
  const keys = Object.keys(examples).filter(
    (functionName) =>
      functionName.startsWith("ex") &&
      functionName.toLowerCase().indexOf("dash") < 0,
  );

  const generateExcel = (
    event: React.MouseEvent<HTMLButtonElement, MouseEvent>,
    data: DataModel.ExcelTable,
  ) => {
    import("mr-excel").then((module) => {
      module.generateExcel(data);
    });
  };
  let items: string[] = Array(32).fill("");

  return (
    <>
      <h3>MR Excel Typescript example</h3>
      <div className="container">
        {items.map((item: string, index: number) => {
          let ex = example[("ex" + (index + 1)) as keyof object];
          return (
            <button
              key={"ex-" + index}
              onClick={(
                event: React.MouseEvent<HTMLButtonElement, MouseEvent>,
              ) => {
                if (typeof ex === "function") {
                  generateExcel(event, (ex as Function)());
                }
              }}
            >
              Example {index + 1}
            </button>
          );
        })}
        <SharedExample/>
      </div>
    </>
  );
}

export default App;
