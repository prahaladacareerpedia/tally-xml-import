import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import './UploadForm.css';

function UploadForm() {
  const [file, setFile] = useState(null);

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (file) {
      const data = await readExcel(file);
      const xmlContent = convertToXML(data);
      const blob = new Blob([xmlContent], { type: 'text/xml' });
      saveAs(blob, 'VoucherData.xml');
    }
  };

  const readExcel = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);
        resolve(json);
      };
      reader.onerror = (err) => reject(err);
      reader.readAsArrayBuffer(file);
    });
  };

  const convertToXML = (data) => {
    let xml = `<?xml version="1.0" encoding="UTF-8"?>\n<ENVELOPE>\n\t<HEADER>\n\t\t<TALLYREQUEST>Import Data</TALLYREQUEST>\n\t</HEADER>\n\t<BODY>\n\t\t<IMPORTDATA>\n\t\t\t<REQUESTDESC>\n\t\t\t\t<REPORTNAME>All Masters</REPORTNAME>\n\t\t\t</REQUESTDESC>\n\t\t\t<REQUESTDATA>\n`;

    data.forEach((row) => {
      const formattedDate = formatDate(row['DATE']);
      
      xml += `\t\t\t\t<TALLYMESSAGE xmlns:UDF="TallyUDF">\n`;
      xml += `\t\t\t\t\t<VOUCHER VCHTYPE="${row['VOUCHER TYPE']}" ACTION="Create">\n`;
      
      xml += `\t\t\t\t\t\t<VOUCHERTYPENAME>${row['VOUCHER TYPE']}</VOUCHERTYPENAME>\n`;
      xml += formattedDate ? `\t\t\t\t\t\t<DATE>${formattedDate}</DATE>\n` : `\t\t\t\t\t\t<DATE></DATE>\n`;
      xml += `\t\t\t\t\t\t<EFFECTIVEDATE>${formattedDate}</EFFECTIVEDATE>\n`;
      
      xml += `\t\t\t\t\t\t<PARTYLEDGERNAME>${row['PARTY NAME'] || ''}</PARTYLEDGERNAME>\n`;
      xml += `\t\t\t\t\t\t<NARRATION>${row['STANDARD NARRATION'] || ''}</NARRATION>\n`;
      xml += `\t\t\t\t\t\t<VOUCHERNUMBER>${row['VOUCHER NUMBER'] || ''}</VOUCHERNUMBER>\n`;
      xml += `\t\t\t\t\t\t<REFERENCE>${row['REFERENCE NUMBER'] || ''}</REFERENCE>\n`;

      // Dynamically add ledger entries
      Object.keys(row).forEach((key) => {
        if (key.startsWith('LEDGER NAME DR/CR') && row[key]) {
          const ledgerIndex = key.match(/\d+/)[0];
          const ledgerName = row[key];
          const effect = row[`EFFECT ${ledgerIndex}`];
          const amount = row[`AMOUNT ${ledgerIndex}`];

          xml += `\t\t\t\t\t\t<ALLLEDGERENTRIES.LIST>\n`;
          xml += `\t\t\t\t\t\t\t<LEDGERNAME>${ledgerName}</LEDGERNAME>\n`;
          xml += `\t\t\t\t\t\t\t<ISDEEMEDPOSITIVE>${effect === 'Debit' ? 'Yes' : 'No'}</ISDEEMEDPOSITIVE>\n`;
          xml += `\t\t\t\t\t\t\t<AMOUNT>${effect === 'Debit' ? '-' : ''}${amount}</AMOUNT>\n`;
          xml += `\t\t\t\t\t\t</ALLLEDGERENTRIES.LIST>\n`;
        }
      });

      xml += `\t\t\t\t\t</VOUCHER>\n`;
      xml += `\t\t\t\t</TALLYMESSAGE>\n`;
    });

    xml += `\t\t\t</REQUESTDATA>\n\t\t</IMPORTDATA>\n\t</BODY>\n</ENVELOPE>`;
    return xml;
  };

  const formatDate = (excelDate) => {
    if (typeof excelDate === 'number') {
      const date = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
      return date.toISOString().split('T')[0].replace(/-/g, '');
    } else if (typeof excelDate === 'string') {
      const date = new Date(excelDate);
      return date.toISOString().split('T')[0].replace(/-/g, '');
    } else {
      return '';
    }
  };

  return (
    <div className="upload-form-container">
      <form onSubmit={handleSubmit} className="upload-form">
        <div className="form-group">
          <label htmlFor="fileInput">Upload Excel File (*.xlsx)</label>
          <input
            type="file"
            accept=".xlsx"
            id="fileInput"
            onChange={handleFileChange}
            required
          />
        </div>
        <div className="generate-btn">
          <button type="submit" disabled={!file}>
            Generate XML
          </button>
        </div>
        <p className="note">
          Note: I have confirm that only 100 entries imported
        </p>
      </form>
    </div>
  );
}

export default UploadForm;
