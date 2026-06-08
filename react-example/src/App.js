import './App.css';
import { generateExcel as excel } from "mr-excel";
import { useEffect, useState } from 'react';
function generateExcel(data) {
  excel(data)
}
function App() {
  const [examples, setExamples] = useState({})
  useEffect(() => {
    import('shared-example').then(examples => {
      setExamples(examples)
    }).catch(e => {
      console.log(e);
    })
  }, [])
  const keys = Object.keys(examples).filter(functionName => functionName.startsWith("ex") && functionName.toLowerCase().indexOf("dash") < 0)

  if (keys.length)
    return <>
      <h3>MR Excel Shared examples</h3>
      <div className='container'>
        {keys.map(functionName => {
          return <button key={functionName} onClick={() => generateExcel(examples[functionName]())}>{functionName}</button>
        })}
      </div></>
  else {
    return <p>loading</p>
  }
}

export default App;
