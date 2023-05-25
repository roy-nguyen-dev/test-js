function parseCSV(csvText) {
    const lines = csvText.split('\n');
    const headers = lines[0].split(',');
    const data = [];
  
    for (let i = 1; i < lines.length; i++) {
      const values = lines[i].split(',');
      if (values.length === headers.length) {
        const row = {};
        for (let j = 0; j < headers.length; j++) {
          row[headers[j]] = values[j];
        }
        data.push(row);
      }
    }
  
    return data;
  }
  
  function readCSVFile(file) {
    fetch(file)
      .then(response => response.text())
      .then(csvText => {
        const data = parseCSV(csvText);
        console.log(data); // Do something with the parsed data
      })
      .catch(error => {
        console.error('Error reading CSV file:', error);
      });
  }
  
  // Usage
  readCSVFile('cmdb_ci1.csv');