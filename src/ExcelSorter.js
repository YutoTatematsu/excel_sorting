import React, { useState } from 'react';
import * as XLSX from 'xlsx';

function ExcelSorter() {
  const [files, setFiles] = useState([]);
  const [sortedFiles, setSortedFiles] = useState([]);
  const [status, setStatus] = useState('');

  const handleDrop = (event) => {
    event.preventDefault();
    const droppedFiles = Array.from(event.dataTransfer.files);
    const excelFiles = droppedFiles.filter(file =>
      (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) &&
      !files.some(existingFile => existingFile.name === file.name)
    );
    setFiles([...files, ...excelFiles]);
  };

  const handleReset = () => {
    setFiles([]);
    setSortedFiles([]);
    setStatus('');
  };

  const handleExecute = async () => {
    const sorted = await Promise.all(files.map(async (file) => {
      const result = await checkExcelFile(file);
      return { name: file.name, ...result };
    }));
    setSortedFiles(sorted);
  };

  const checkExcelFile = (file) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(worksheet['!ref']);

        let dates = [];
        let boardingStations = [];
        let alightingStations = [];
        let tripTypes = [];
        let expenseTypes = [];
        let destinations = [];
        let transportTypes = [];
        let amounts = [];
        let status = 'OK';
        let commutingEntries = [];

        //データの取得
        for (let row = range.s.r; row <= range.e.r; row++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: 1 }); // B列のセル（B1という表示）
          const cell = worksheet[cellAddress]; // B列のセルの記入内容

          // //B列が存在する || B列の値が1以上かどうか
          // console.log((cell ? cell.v : "null") + " " + cellAddress);

          //記入内容がある && 値が1以上
          if (cell && cell.v >= 1) {

            //データを取得
            const cData = worksheet[XLSX.utils.encode_cell({ r: row, c: 2 })]?.v;// C列
            const dData = worksheet[XLSX.utils.encode_cell({ r: row, c: 3 })]?.v;// D列
            const fData = worksheet[XLSX.utils.encode_cell({ r: row, c: 5 })]?.v;// F列
            const gData = worksheet[XLSX.utils.encode_cell({ r: row, c: 6 })]?.v;// G列
            const hData = worksheet[XLSX.utils.encode_cell({ r: row, c: 7 })]?.v;// H列
            const iData = worksheet[XLSX.utils.encode_cell({ r: row, c: 8 })]?.v;// I列
            const jData = worksheet[XLSX.utils.encode_cell({ r: row, c: 9 })]?.v;// J列
            const kData = worksheet[XLSX.utils.encode_cell({ r: row, c: 10 })]?.v;// K列

            //データが記入されているか判定
            if (!cData || !dData || !fData || !gData || !hData || !iData || !jData || !kData) {
              console.log(cellAddress + "にて終了" + cData + " " + dData + " " + fData + " " + gData + " " + hData + " " + iData + " " + jData + " " + kData);
              break;
            }

            //データを追加
            dates.push(cData);
            boardingStations.push(dData);
            alightingStations.push(fData);
            tripTypes.push(gData);
            expenseTypes.push(hData);
            destinations.push(iData);
            transportTypes.push(jData);
            amounts.push(kData);

            //通勤費の判定に使うので別途記録
            if (hData === '通勤費') {
              //追加
              commutingEntries.push({
                date: cData,
                boardingStation: dData,
                alightingStation: fData,
                tripType: gData,
                expenseType: hData,
                destination: iData,
                transportType: jData,
                amount: kData
              });
            }
          }
        }

        //通勤費の判定（現状：1行で収まるかつ同じ区間である場合のみ　複数未対応「バス」「電車」）
        if (commutingEntries.length >= Math.floor(expenseTypes.length / 2) &&//出勤日の半数か判定
          commutingEntries.every(entry => entry.boardingStation === commutingEntries[0].boardingStation)//すべてのデータが同じか判定
        ) {
          //完全に同じ経路を使用している
          status = 'OK';
        } else {
          status = '問題あり';
        }

        resolve({
          status,
          dates,
          boardingStations,
          alightingStations,
          tripTypes,
          expenseTypes,
          destinations,
          transportTypes,
          amounts
        });
      };
      reader.readAsArrayBuffer(file);
    });
  };

  const handleRemoveFile = (index) => {
    const newFiles = [...files];
    newFiles.splice(index, 1);
    setFiles(newFiles);
  };

  return (
    <div>
      <div
        onDrop={handleDrop}
        onDragOver={(event) => event.preventDefault()}
        style={{ width: '300px', height: '300px', border: '1px dashed #ccc' }}
      >
        エクセルファイルをドロップしてください
      </div>
      {files.length > 0 && (
        <div>
          <h2>ドロップされたファイル:</h2>
          <ul>
            {files.map((file, index) => (
              <li key={index}>
                {file.name}
                <button onClick={() => handleRemoveFile(index)}>削除</button>
              </li>
            ))}
          </ul>
        </div>
      )}
      <button onClick={handleReset}>リセット</button>
      <button onClick={handleExecute}>実行</button>
      {sortedFiles.length > 0 && (
        <div>
          {sortedFiles.map((file, fileIndex) => (
            <div key={fileIndex}>
              <h3>{file.name} - {file.status}</h3>
              <table>
                <thead>
                  <tr>
                    <th>Date</th>
                    <th>Boarding Station</th>
                    <th>Alighting Station</th>
                    <th>Trip Type</th>
                    <th>Expense Type</th>
                    <th>Destination</th>
                    <th>Transport Type</th>
                    <th>Amount</th>
                  </tr>
                </thead>
                <tbody>
                  {file.dates.map((_, entryIndex) => (
                    <tr key={entryIndex}>
                      <td>{file.dates[entryIndex]}</td>
                      <td>{file.boardingStations[entryIndex]}</td>
                      <td>{file.alightingStations[entryIndex]}</td>
                      <td>{file.tripTypes[entryIndex]}</td>
                      <td>{file.expenseTypes[entryIndex]}</td>
                      <td>{file.destinations[entryIndex]}</td>
                      <td>{file.transportTypes[entryIndex]}</td>
                      <td>{file.amounts[entryIndex]}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default ExcelSorter;
