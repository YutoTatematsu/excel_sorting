import React, { useState } from 'react';
import * as XLSX from 'xlsx'; //エクセルを使うためのインポート
import './ExcelSorter.css';   //CSSをインポート

function ExcelSorter() {
  const [files, setFiles] = useState([]);
  const [sortedFiles, setSortedFiles] = useState([]);
  const [filterStatus, setFilterStatus] = useState('all');
  const [openFiles, setOpenFiles] = useState({});

  /**ファイルドロップ時の処理
   * @param {*} event 
   */
  const handleDrop = (event) => {
    // イベントのデフォルト動作をキャンセル（必要？）
    event.preventDefault();
    // ドロップされたファイルの一覧を取得し配列に代入
    const droppedFiles = Array.from(event.dataTransfer.files);
    // エクセルかどうか判定 && 同名ファイルでない判定
    const excelFiles = droppedFiles.filter(file =>
      (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) &&
      !files.some(existingFile => existingFile.name === file.name)
    );

    /** スプレッド演算子は、配列やオブジェクトを展開するための構文（...変数名）
      * 配列に使うと、配列内の各要素を個別の要素として展開する
      * 配列内の要素を新しい配列に含める
      * 例[...array]とすることで、array内の要素を展開して新しい配列に追加
      */
    setFiles([...files, ...excelFiles]);
  };

  /**リセットボタンを押したときの処理
   */
  const handleReset = () => {
    // 配列などの読み込んだデータを初期化
    setFiles([]);
    setSortedFiles([]);
    setFilterStatus('all');
    setOpenFiles({});
  };

  /** 実行ボタンを押したときの処理
   */
  const handleExecute = async () => {
    // 非同期で行うためPromise.allすべての非同期処理が完了するのを待つ
    const sorted = await Promise.all(files.map(async (file) => {// 配列の各要素に対して非同期の関数を実行
      // ファイルの読み込み＆処理関数を起動
      const result = await checkExcelFile(file);
      // 上の結果を result に格納
      return { name: file.name, ...result };
    }));
    // 配列をコンポーネントの状態に設定
    setSortedFiles(sorted);
  };

  const checkExcelFile = (file) => {
    // 新しいPromiseオブジェクトを作成、resolve関数はPromiseが成功したときに結果を返す
    return new Promise((resolve) => {
      // ファイル読み込みを作成
      const reader = new FileReader();

      // ファイルが読み込まれたときに行う処理を定義
      reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);         //読み込んだファイルデータを取得（バイナリ）
        const workbook = XLSX.read(data, { type: 'array' });      //XLSXライブラリを使用：上データをExcelファイルにする
        const sheetName = workbook.SheetNames[0];                 //最初のシート名を取得
        const worksheet = workbook.Sheets[sheetName];             //シート名を使用し対応するデータを取得
        const range = XLSX.utils.decode_range(worksheet['!ref']); //ワークシートのデータ範囲を解析、データが含まれるセルの範囲を取得

        let dates = [];             //日付
        let boardingStations = [];  //乗車駅
        let alightingStations = []; //降車駅
        let tripTypes = [];         //片道・往復
        let expenseTypes = [];      //通勤費か業務交通費
        let destinations = [];      //目的地
        let transportTypes = [];    //通勤手段
        let amounts = [];           //金額
        let status = 'OK';          //判定結果
        let commutingEntries = [];  //通勤費すべてのデータ
        let typeEntries = [[], []]; //経路データ

        //データを取得
        for (let row = range.s.r; row <= range.e.r; row++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: 1 });//B列のセル
          const cell = worksheet[cellAddress];//今見ているB列の中身

          //B列の中身が数値かつ1以上なら
          if (cell && cell.v >= 1) {
            //データを取得
            const c_Date = worksheet[XLSX.utils.encode_cell({ r: row, c: 2 })]?.v;//日付
            const d_SStation = worksheet[XLSX.utils.encode_cell({ r: row, c: 3 })]?.v;//乗車駅
            const f_EStation = worksheet[XLSX.utils.encode_cell({ r: row, c: 5 })]?.v;//降車駅
            const g_TripType = worksheet[XLSX.utils.encode_cell({ r: row, c: 6 })]?.v;//片道・往復
            const h_ExpenseTypes = worksheet[XLSX.utils.encode_cell({ r: row, c: 7 })]?.v;//通勤費か業務交通費
            const i_Destinations = worksheet[XLSX.utils.encode_cell({ r: row, c: 8 })]?.v;//目的地
            const j_TransportType = worksheet[XLSX.utils.encode_cell({ r: row, c: 9 })]?.v;//通勤手段
            const k_Money = worksheet[XLSX.utils.encode_cell({ r: row, c: 10 })]?.v;//金額

            //データがなかった場合（記入漏れ判別未導入）
            if (!c_Date || !d_SStation || !f_EStation || !g_TripType || !h_ExpenseTypes || !i_Destinations || !j_TransportType || !k_Money) {
              //読み込み終了
              break;
            }

            //データ記録配列に追加
            dates.push(c_Date);
            boardingStations.push(d_SStation);
            alightingStations.push(f_EStation);
            tripTypes.push(g_TripType);
            expenseTypes.push(h_ExpenseTypes);
            destinations.push(i_Destinations);
            transportTypes.push(j_TransportType);
            amounts.push(k_Money);

            //通勤費の可否判定のために別途記録
            if (h_ExpenseTypes === '通勤費') {

              let found = false;//データを追加できたか

              for (let i = 0; i < typeEntries.length; i++) {
                if (typeEntries[i].length === 0 ||//初記録
                  (typeEntries[i][0].transportType === j_TransportType &&
                    typeEntries[i][0].boardingStation === d_SStation &&
                    typeEntries[i][0].alightingStation === f_EStation)
                ) {
                  //追加
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

                  //追加したのでフラグをtrue
                  found = true;
                  break;
                }
              }

              //追加できなかった場合
              if (!found) {
                //経路を増やしデータを追加する
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

              //通勤費すべてのデータを記録
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

        /* 交通費に分類される経路の判定 */
        //出勤日数の半分以上が通勤費になっているか(通勤ルートで割ることで日数を出す)
        if ((commutingEntries.length / typeEntries.length) >= (commutingEntries.length / typeEntries.length) / 2) {

          //同じ経路内で要素の変化がないか判定（「金額」「片道・往復」「目的地」「交通機関」が区間内で同じか）
          if (typeEntries.every(route => allRouteEqualCheck(route))) {

            //金額が大きすぎたり && 低すぎないか判定
            if (commutingEntries.every(route => route.amount <= 1000) &&
              commutingEntries.every(route => route.amount >= 410)) {
              status = 'OK';
            }
            else {
              console.log("値段の大きさが水準から出ています");
              status = '注意';
            }
          } else {
            console.log("同じ経路内で異なる項目があります");
            status = '問題あり';
          }
        } else {
          console.log("出勤日数が半分以上の制約を満たしていません");
          status = '問題あり';
        }

        // リゾルブで非同期処理のデータを返す（resolveはPromiseを完了させ、その結果を返すための関数とのこと）
        // returnのようなもの
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

  /**元の配列を触らずに表示するファイルを変更する
   * @param {削除する要素番号} index 
   */
  const handleRemoveFile = (index) => {
    // 配列をコピーします（スプレッド演算子を使用して新しい配列を作成）
    const newFiles = [...files];
    // 指定インデックスの要素を削除
    newFiles.splice(index, 1);
    // 更新ファイル配列を表示させる（再レンダリングされ画面表示変わる）
    setFiles(newFiles);
  };

  /**中身の要素がすべて同じならtrue
   * @param {確認したい配列} array 
   * @returns 
   */
  const allRouteEqualCheck = (array) => {
    // すべてのデータのすべての要素を比較
    if (array.every(entry => entry.boardingStation === array[0].boardingStation) &&
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

  /**フィルター処理
   * 表示するファイルを設定
   */
  const filteredFiles = sortedFiles.filter(file => {
    // ページ上で入力したものを判定
    if (filterStatus === 'all') return true;

    // 表示したいステータスファイルを返す
    return file.status === filterStatus;
  });

  /**判定後の要素がクリックされたとき
   * @param {クリックされた要素番号} index 
   */
  const toggleFileOpen = (index) => {
    // 状態更新用の変数を書き換え
    setOpenFiles(prevState => ({
      // 前の状態のprevStateを受け取り新しい状態を返す
      ...prevState,
      // 開閉フラグを反転させる
      [index]: !prevState[index]
    }));
  };

  /**日付表示処理
   * @param {エクセルデータ} excelDate 
   * @returns 
   */
  const formatDate = (excelDate) => {
    // 日付を直す
    const jsDate = new Date((excelDate - 25569) * 86400 * 1000);
    // 月
    const month = jsDate.getMonth() + 1;
    // 日
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
