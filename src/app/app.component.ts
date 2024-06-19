import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})
export class AppComponent {
  title = 'testAng';

  // downloadCSV() {
  //   // Define the data
  //   const data = [
  //     { name: 'John Doe', age: 28, city: 'New York' },
  //     { name: 'Anna Smith', age: 24, city: 'London' },
  //     { name: 'Peter Jones', age: 45, city: 'Paris' }
  //   ];

  //   // Convert data to CSV format
  //   const csvData = this.convertToCSV(data);

  //   // Create a Blob from the CSV data
  //   const blob = new Blob([csvData], { type: 'text/csv' });

  //   // Create a link element
  //   const link = document.createElement('a');

  //   // Set the href attribute to a URL created from the Blob
  //   const url = window.URL.createObjectURL(blob);
  //   link.href = url;

  //   // Set the download attribute to the desired file name
  //   link.download = 'data.csv';

  //   // Append the link to the document body
  //   document.body.appendChild(link);

  //   // Programmatically click the link to trigger the download
  //   link.click();

  //   // Clean up by removing the link and revoking the Object URL
  //   document.body.removeChild(link);
  //   window.URL.revokeObjectURL(url);
  // }

  // convertToCSV(objArray: any[]): string {
  //   const array = typeof objArray !== 'object' ? JSON.parse(objArray) : objArray;
  //   let str = '';
  //   const header = Object.keys(array[0]).join(',') + '\r\n';
  //   str += header;

  //   for (const row of array) {
  //     let line = '';
  //     for (const index in row) {
  //       if (line !== '') line += ',';
  //       line += row[index];
  //     }
  //     str += line + '\r\n';
  //   }
  //   return str;
  // }


  downloadExcel() {
    // Define the data
    const data = [
      { name: 'John Doe', age: 28, city: 'New York' },
      { name: 'Anna Smith', age: 24, city: 'London' },
      { name: 'Peter Jones', age: 45, city: 'Paris' }
    ];

    // Create a new workbook
    const workbook: XLSX.WorkBook = XLSX.utils.book_new();

    // Convert data to sheet
    const sheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);

    // Add the sheet to the workbook
    XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');

    // Generate a Blob object for the workbook
    const wbout: ArrayBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    // Create a Blob object
    const blob = new Blob([wbout], { type: 'application/octet-stream' });

    // Create a link element
    const link = document.createElement('a');

    // Set the href attribute to a URL created from the Blob
    const url = window.URL.createObjectURL(blob);
    link.href = url;

    // Set the download attribute to the desired file name
    link.download = 'data.xlsx';

    // Append the link to the document body
    document.body.appendChild(link);

    // Programmatically click the link to trigger the download
    link.click();

    // Clean up by removing the link and revoking the Object URL
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
  }
}
