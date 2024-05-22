import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './ExcelSorter.css';

function ExcelSorter() {
  const [files, setFiles] = useState([]);
  const [sortedFiles, setSortedFiles] = useState([]);
  const [filterStatus, setFilterStatus] = useState('all');
  const [openFiles, setOpenFiles] = useState({});

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
    setFilterStatus('all');
    setOpenFiles({});
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
        let typeEntries = [[], []];

        for (let row = range.s.r; row <= range.e.r; row++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: 1 });
          const cell = worksheet[cellAddress];

          if (cell && cell.v >= 1) {
            const c_Date = worksheet[XLSX.utils.encode_cell({ r: row, c: 2 })]?.v;
            const d_SStation = worksheet[XLSX.utils.encode_cell({ r: row, c: 3 })]?.v;
            const f_EStation = worksheet[XLSX.utils.encode_cell({ r: row, c: 5 })]?.v;
            const g_TripType = worksheet[XLSX.utils.encode_cell({ r: row, c: 6 })]?.v;
            const h_ExpenseTypes = worksheet[XLSX.utils.encode_cell({ r: row, c: 7 })]?.v;
            const i_Destinations = worksheet[XLSX.utils.encode_cell({ r: row, c: 8 })]?.v;
            const j_TransportType = worksheet[XLSX.utils.encode_cell({ r: row, c: 9 })]?.v;
            const k_Money = worksheet[XLSX.utils.encode_cell({ r: row, c: 10 })]?.v;

            if (!c_Date || !d_SStation || !f_EStation || !g_TripType || !h_ExpenseTypes || !i_Destinations || !j_TransportType || !k_Money) {
              break;
            }

            dates.push(c_Date);
            boardingStations.push(d_SStation);
            alightingStations.push(f_EStation);
            tripTypes.push(g_TripType);
            expenseTypes.push(h_ExpenseTypes);
            destinations.push(i_Destinations);
            transportTypes.push(j_TransportType);
            amounts.push(k_Money);

            if (h_ExpenseTypes === '通勤費') {
              let found = false;

              for (let i = 0; i < typeEntries.length; i++) {
                if (
                  typeEntries[i].length === 0 ||
                  (typeEntries[i][0].transportType === j_TransportType &&
                    typeEntries[i][0].boardingStation === d_SStation &&
                    typeEntries[i][0].alightingStation === f_EStation)
                ) {
                  typeEntries[i].push({
                    date: c_Date,
                    boardingStation: d_SStation,
                    alightingStation: f_EStation,
                    tripType: g_TripType,
                    expenseType: h_ExpenseTypes,
                    destination: i_Destinations,
                    transportType: j_TransportType,
                    amount: k_Money
                  });
                  found = true;
                  break;
                }
              }

              if (!found) {
                typeEntries.push([{
                  date: c_Date,
                  boardingStation: d_SStation,
                  alightingStation: f_EStation,
                  tripType: g_TripType,
                  expenseType: h_ExpenseTypes,
                  destination: i_Destinations,
                  transportType: j_TransportType,
                  amount: k_Money
                }]);
              }

              commutingEntries.push({
                date: c_Date,
                boardingStation: d_SStation,
                alightingStation: f_EStation,
                tripType: g_TripType,
                expenseType: h_ExpenseTypes,
                destination: i_Destinations,
                transportType: j_TransportType,
                amount: k_Money
              });
            }
          }
        }

        if (commutingEntries.length >= Math.floor(expenseTypes.length / 2)) {
          const allRoutesValid = typeEntries.every(route => allRouteEqualCheck(route));
          if (allRoutesValid) {
            status = 'OK';
          } else {
            status = '注意';
          }
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

  const allRouteEqualCheck = (array) => {
    if (
      array.every(entry => entry.boardingStation === array[0].boardingStation) &&
      array.every(entry => entry.alightingStation === array[0].alightingStation) &&
      array.every(entry => entry.tripType === array[0].tripType) &&
      array.every(entry => entry.expenseType === array[0].expenseType) &&
      array.every(entry => entry.destination === array[0].destination) &&
      array.every(entry => entry.transportType === array[0].transportType) &&
      array.every(entry => entry.amount === array[0].amount)
    ) {
      return true;
    }
    return false;
  };

  const filteredFiles = sortedFiles.filter(file => {
    if (filterStatus === 'all') return true;
    return file.status === filterStatus;
  });

  const toggleFileOpen = (index) => {
    setOpenFiles(prevState => ({
      ...prevState,
      [index]: !prevState[index]
    }));
  };

  const formatDate = (excelDate) => {
    const jsDate = new Date((excelDate - 25569) * 86400 * 1000);
    const month = jsDate.getMonth() + 1;
    const day = jsDate.getDate();
    return `${month}/${day}`;
  };

  return (
    <div className="container">
      <div className="dropzone" onDragOver={(e) => e.preventDefault()} onDrop={handleDrop}>
        ドラッグ & ドロップでファイルを追加
      </div>
      <table className="file-list">
        <thead>
          <tr>
            <th>ファイル名</th>
            <th>操作</th>
          </tr>
        </thead>
        <tbody>
          {files.map((file, index) => (
            <tr key={index} className="file-item">
              <td className="file-name">{file.name}</td>
              <td>
                <button className="remove-button" onClick={() => handleRemoveFile(index)}>削除</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      <div className="controls">
        <button className="control-button" onClick={handleReset}>リセット</button>
        <button className="control-button" onClick={handleExecute}>実行</button>
      </div>
      <div className="filter">
        <label>フィルタリング:
          <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
            <option value="all">すべて</option>
            <option value="OK">OK</option>
            <option value="注意">注意</option>
            <option value="問題あり">問題あり</option>
          </select>
        </label>
      </div>
      {sortedFiles.length > 0 && (
        <div className="sorted-files">
          {filteredFiles.map((file, index) => (
            <div key={index} className="sorted-file">
              <div className="file-header" onClick={() => toggleFileOpen(index)}>
                <span className="file-name">{file.name}</span>
                <span className="file-status">{file.status}</span>
                <span>{openFiles[index] ? '▲' : '▼'}</span>
              </div>
              {openFiles[index] && (
                <table className="file-table">
                  <thead>
                    <tr>
                      <th>日付</th>
                      <th>乗車駅</th>
                      <th>降車駅</th>
                      <th>片道・往復</th>
                      <th>通勤・業務</th>
                      <th>目的地</th>
                      <th>交通機関種類</th>
                      <th>金額</th>
                    </tr>
                  </thead>
                  <tbody>
                    {file.dates.map((date, entryIndex) => (
                      <tr key={entryIndex}>
                        <td>{formatDate(date)}</td>
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
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default ExcelSorter;
