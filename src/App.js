import React, { Component } from 'react';
import logo from './logo.svg';
import './App.css';
import XLSX from 'xlsx';

class App extends Component {
  constructor(props) {
    super(props)
    this.state = { excelData: [] }
  }


  handleChange(e) {
    e.stopPropagation(); e.preventDefault();

    let files = e.target.files, f = files[0];
    let reader = new FileReader();

    reader.onload = (e) => {
      let data = e.target.result;
      if (!rABS) data = new Uint8Array(data);
      let workbook = XLSX.read(data, { type: rABS ? 'binary' : 'array' });

      /* DO SOMETHING WITH workbook HERE */
      let worksheet = workbook.Sheets[workbook.SheetNames[0]];
      let container = document.getElementById('tableau');
      let json = XLSX.utils.sheet_to_json(worksheet);
      this.setState({ excelData: json });
      // var html = XLSX.utils.sheet_to_html(worksheet);
      // container.innerHTML = html;
    };

    let rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
    if (rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);

  }

  renderRow = (items) => {
    return (
      items.map(function (item, index) {
        return (
          <tr key={index}>
            <td>{item.OrderDate}</td>
            <td>{item.Region}</td>
            <td>{item.Rep}</td>
            <td>{item.Item}</td>
            <td>{item.Units}</td>
            <td>{index}</td>
          </tr>
        );
      })
    )
  };

  downloadExcel = (e) => {
    var filename = "report.xlsx";
    var data = [[1, 1, 1], [2, 2, 2], [3, 3, 3]];
    var ws_name = "Sheet1";

    // if (typeof console !== 'undefined') console.log(new Date());
    var wb = XLSX.utils.book_new(), ws = XLSX.utils.aoa_to_sheet(data);

    /* add worksheet to workbook */
    XLSX.utils.book_append_sheet(wb, ws, ws_name);

    /* write workbook */
    // if (typeof console !== 'undefined') console.log(new Date());
    XLSX.writeFile(wb, filename);
    // if (typeof console !== 'undefined') console.log(new Date());
  }

  render() {
    return (
      <div className="App" >
        <input type="file"
          name="myFile"
          onChange={(e) => this.handleChange(e)} />
        <br /><br />
        <table className={"table table-hover"}>
          <thead>
            <tr>
              <th>Order Date</th>
              <th>Region</th>
              <th>Rep</th>
              <th>Item</th>
              <th>Units</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {
              this.renderRow(this.state.excelData)
            }
            {
              this.state.excelData.length === 0 && (<tr><td colSpan={5}>No record(s) found.</td></tr>)
            }
          </tbody>
        </table>
        <hr />
        <button onClick={e => this.downloadExcel()}>Download Excel</button>
      </div>
    );
  }
}

export default App;
