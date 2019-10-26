import React from 'react';
import './App.css';
import FileSaver from 'file-saver'

const Excel = require("exceljs");

// Create workbook & add worksheet
const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet("ExampleSheet");

class App extends React.Component {

  ExcelJS = () => {

    // add image to workbook by filename
    var imageId1 = workbook.addImage({
      filename: '../public/logo192.png',
      extension: 'jpeg',
    });

    // add column headers
    worksheet.columns = [
      { header: 'Package', key: 'package_name' },
      { header: 'Author', key: 'author_name' }
    ];

    // Add row using key mapping to columns
    worksheet.addRow(
      { package_name: "ABC", author_name: "Author 1" },
      { package_name: "XYZ", author_name: "Author 2" }
    );

    // Add rows as Array values
    worksheet.addRow(["BCD", "Author Name 3"], [{imageId1}, "Author Name 3"]);

    // Add rows using both the above of rows
    const rows = [
      ["FGH", "Author Name 4"],
      { package_name: "PQR", author_name: "Author 5" }
    ];

    worksheet
      .addRows(rows);

    // save workbook to disk
    // workbook.xlsx.writeFile('sample.xlsx')
    //   .then(() => {
    //     console.log("saved");
    //   })
    //   .catch((err) => {
    //     console.log("err", err);
    //   });
   workbook.xlsx.writeBuffer()
  .then(buffer => {
    FileSaver.saveAs(new Blob([buffer]), `${Date.now()}_feedback.xlsx`);
    console.log('Berhasil Yeay')
  })
  .catch(err => console.log('Error writing excel export', err))
    
  }

  render () {
    return (
      <div>
        Excel JS : 
        <button onClick={this.ExcelJS}>Download as XLSX</button>
      </div>
    );
  }
}

export default App;
