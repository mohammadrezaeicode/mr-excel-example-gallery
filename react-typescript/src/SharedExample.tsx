import { generateExcel as excel } from "mr-excel";
import { useEffect, useState } from 'react';
function generateExcel(data:any) {
  excel(data)
}
function SharedExample() {
  const [examples, setExamples] = useState<any>({});
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

  if (keys.length)
    return (
      <>
          {keys.map((functionName) => {
            return (
              <button
                key={functionName}
                onClick={() => generateExcel(examples[functionName]())}
              >
                {functionName}
              </button>
            );
          })}
      </>
    );
  else {
    return <p>loading</p>;
  }
}

export default SharedExample;
