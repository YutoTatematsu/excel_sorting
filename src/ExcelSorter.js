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

        //データ配列
        let dates = [];
        let boardingStations = [];
        let alightingStations = [];
        let tripTypes = [];
        let expenseTypes = [];
        let destinations = [];
        let transportTypes = [];
        let amounts = [];
        let status = 'OK';
        let commutingEntries = [];//通勤費をまとめる変数
        let typeEntries = [,];    //通勤費の交通経路を記録する2次元配列
        let firstTypeEntries = [];//1つ目の交通手段を記録する変数
        let secondTypeEntries = [];//2つ目の交通手段を記録する変数

        //データの取得
        for (let row = range.s.r; row <= range.e.r; row++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: 1 }); // B列のセル（B1という表示）
          const cell = worksheet[cellAddress]; // B列のセルの記入内容

          // //B列が存在する || B列の値が1以上かどうか
          // console.log((cell ? cell.v : "null") + " " + cellAddress);

          //記入内容がある && 値が1以上
          if (cell && cell.v >= 1) {

            //データを取得
            const c_Date = worksheet[XLSX.utils.encode_cell({ r: row, c: 2 })]?.v;// C列（日付）
            const d_SStation = worksheet[XLSX.utils.encode_cell({ r: row, c: 3 })]?.v;// D列（乗車駅名）
            const f_EStation = worksheet[XLSX.utils.encode_cell({ r: row, c: 5 })]?.v;// F列（降車駅名）
            const g_TripType = worksheet[XLSX.utils.encode_cell({ r: row, c: 6 })]?.v;// G列（片道・往復）
            const h_ExpenseTypes = worksheet[XLSX.utils.encode_cell({ r: row, c: 7 })]?.v;// H列（通勤・業務）
            const i_Destinations = worksheet[XLSX.utils.encode_cell({ r: row, c: 8 })]?.v;// I列（目的地）
            const j_TransportType = worksheet[XLSX.utils.encode_cell({ r: row, c: 9 })]?.v;// J列（使用交通機関）
            const k_Money = worksheet[XLSX.utils.encode_cell({ r: row, c: 10 })]?.v;// K列（金額）

            //データが記入されているか判定
            if (!c_Date || !d_SStation || !f_EStation || !g_TripType || !h_ExpenseTypes || !i_Destinations || !j_TransportType || !k_Money) {
              // console.log(cellAddress + "にて終了" + c_Date + " " + d_SStation + " " + f_EStation + " " + g_TripType + " " + h_ExpenseTypes + " " + i_Destinations + " " + j_TransportType + " " + k_Money);
              break;
            }

            //データを追加
            dates.push(c_Date);
            boardingStations.push(d_SStation);
            alightingStations.push(f_EStation);
            tripTypes.push(g_TripType);
            expenseTypes.push(h_ExpenseTypes);
            destinations.push(i_Destinations);
            transportTypes.push(j_TransportType);
            amounts.push(k_Money);

            //通勤費の判定に使うので別途記録
            if (h_ExpenseTypes === '通勤費') {

              // if (firstTypeEntries.length >= 1) {
              //   console.log(`${firstTypeEntries[0].transportType} ${j_TransportType}`);
              //   console.log(`${firstTypeEntries[0].boardingStation} ${d_SStation}`);
              //   console.log(`${firstTypeEntries[0].alightingStations} ${f_EStation}`);
              //   console.log(`${firstTypeEntries.length}`);

              //   Debug_AllConsoleLog(firstTypeEntries);
              // }

              //記録するか判定
              if (firstTypeEntries.length === 0 ||                      //初記録
                firstTypeEntries[0].transportType === j_TransportType &&//交通機関が同じ
                firstTypeEntries[0].boardingStation === d_SStation &&  //乗車駅が同じ
                firstTypeEntries[0].alightingStation === f_EStation)   //降車駅が同じ
              {
                //1つ目の交通経路を記録
                firstTypeEntries.push({
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
              //記録するか判定
              else if (secondTypeEntries.length === 0 ||                      //初記録
                secondTypeEntries[0].transportType === j_TransportType &&//交通機関が同じ
                secondTypeEntries[0].boardingStation === d_SStation &&  //乗車駅が同じ
                secondTypeEntries[0].alightingStation === f_EStation)   //降車駅が同じ
              {
                //1つ目の交通経路を記録
                secondTypeEntries.push({
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

              //通勤費すべての配列に追加
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

        /* 通勤費判定 */
        //通勤費の判定（現状：1行で収まるかつ同じ区間である場合のみ　複数未対応「バス」「電車」）
        //出勤日の半数か判定
        if (commutingEntries.length >= Math.floor(expenseTypes.length / 2)) {

          //デバッグ用処理
          // Debug_AllConsoleLog(commutingEntries);
          // Debug_AllConsoleLog(firstTypeEntries);
          // Debug_AllConsoleLog(secondTypeEntries);

          //1つの交通機関を用いて出勤している || 2つの交通機関を用いる場合
          if (allRouteEqualCheck(commutingEntries)
            || (allRouteEqualCheck(firstTypeEntries) && allRouteEqualCheck(secondTypeEntries))
          ) {
            status = 'OK';
          }
          else {
            status = '問題あり';
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

  /** 配列内のすべての項目が同じ場合trueを返す
   * @param {通勤ルート配列} array 
   * @returns 
   */
  const allRouteEqualCheck = (array) => {
    if (
      array.every(entry => entry.boardingStation === array[0].boardingStation) &&
      array.every(entry => entry.alightingStation === array[0].alightingStation) &&
      array.every(entry => entry.tripType === array[0].tripType) &&
      array.every(entry => entry.expenseType === array[0].expenseType) &&
      array.every(entry => entry.destination === array[0].destination) &&
      array.every(entry => entry.transportType === array[0].transportType) &&
      array.every(entry => entry.amount === array[0].amount)) {
      return true;
    }
    return false;
  }


  /**すべてのルートデータを表示
   * @param {通勤ルート配列} array 
   */
  const Debug_AllConsoleLog = (array) => {
    //デバッグ用処理
    array.forEach(entry => console.log(`乗った駅：${entry.boardingStation}`));
    array.forEach(entry => console.log(`降りた駅：${entry.alightingStation}`));
    array.forEach(entry => console.log(`${entry.tripType}`));
    array.forEach(entry => console.log(`${entry.expenseType}`));
    array.forEach(entry => console.log(`目的地：${entry.destination}`));
    array.forEach(entry => console.log(`使用交通機関：${entry.transportType}`));
    array.forEach(entry => console.log(`${entry.amount}円`));
    console.log("配列数" + array.length);
  }

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
