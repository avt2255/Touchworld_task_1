const express = require('express');
const XLSX = require('xlsx');
const mysql = require('mysql');
const multer = require('multer');
const app = express();
//select destination of file
const upload = multer({ dest: 'uploads/' });

//connecting to database in the localhost
const connection = mysql.createConnection({
  host: 'localhost',
  user: 'root',
  password: '',
  database: 'touchworld',
});

connection.connect((err) => {
  if (err) throw err;
  console.log('Connected to MySQL database!');
});

//creating details_tb mysql table
app.get('/createDetailsTable', (req, res) => {
  const createEmployeeTable = `
  CREATE TABLE details_tb (
    Id INT PRIMARY KEY AUTO_INCREMENT,
    name VARCHAR(255),
    address VARCHAR(255),
    age VARCHAR(20),
    date_of_birth VARCHAR(20)
  )
`
  // age and date of birth is taken as varchar .It is because when we take as int and date format if we put values as invalid format that values will automatically convert to 0 and 0000-00-00 format.So we cannot check conditions properly while uploading excel file data to table

  connection.query(createEmployeeTable, (err) => {
    if (err) {
      res.send("Details table already exists")
    } else {
      res.send("Details table created")
    }
  });
});

//creating details2_tb mysql table for storing fetched data from first table.
app.get('/createSecondDetailsTable', (req, res) => {
  const createEmployeeTable = `
  CREATE TABLE details2_tb (
    Id INT PRIMARY KEY AUTO_INCREMENT,
    name VARCHAR(255),
    address VARCHAR(255),
    age VARCHAR(20),
    date_of_birth VARCHAR(20)
  )
`
  connection.query(createEmployeeTable, (err) => {
    if (err) {
      res.send("second table already exists")
    } else {
      res.send("second table created")
    }
  });
});

//inserting some sample data to mysql details_tb
app.get('/insertSampleData', (req, res) => {
  const personalData = [
    { name: 'Arun', address: 'nileswaram P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },
    { name: 'Das', address: 'iritty P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },
    { name: 'Vishnu', address: 'kochi P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },
    { name: 'Manu', address: 'kannur P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },
    { name: 'Sohan', address: 'trikaripur P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },
    { name: 'Aswin', address: 'mangad P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },

  ];

  const insertQuery = 'INSERT INTO details_tb (name, address, age, date_of_birth) VALUES ?';
  const values = personalData.map(employee => [employee.name, employee.address, employee.age, employee.date_of_birth]);

  connection.query(insertQuery, [values], (err) => {
    if (err) {
      console.error('Error inserting sample employees: ', err);
      return res.status(500).json({ error: 'Error inserting sample employees' });
    }
    console.log('Sample employees inserted!');
    return res.status(200).json({ message: 'Sample employees inserted successfully' });
  });
});

//export table data to excel file
app.get('/export', async (req, res) => {
  try {

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet([

      { name: 'Arun', address: 'nileswaram P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },
      { name: 'Das', address: 'iritty P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },
      { name: 'Vishnu', address: 'kochi P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },
      { name: 'Manu', address: 'kannur P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },
      { name: 'Sohan', address: 'trikaripur P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },
      { name: 'Aswin', address: 'mangad P.o chalakkode, kerala,670307,Kannur District', age: '45', date_of_birth: '2023-05-10' },
      // you can add more values as object from the database
    ]);

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Users');
    const excelFilePath = 'users.xlsx';
    XLSX.writeFile(workbook, excelFilePath);

    console.log('Excel file generated successfully:', excelFilePath);
    res.status(200).send('Excel file generated successfully');

  } catch (error) {
    console.error('Error fetching data from MySQL:', error);
    res.status(500).send('Internal server error');
  }
});

// uploading excel file data to another table.We cannot add duplicate data(Id) in to same table.So choose another table 

app.post('/upload', upload.single('file'), (req, res) => {
  const workbook = XLSX.readFile(req.file.path);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = XLSX.utils.sheet_to_json(worksheet);

  const errors = [];
  const validRecords = [];

  for (let i = 0; i < jsonData.length; i++) {
    const record = jsonData[i];

    const id = record.Id;
    const name = record.name;
    const address = record.address;
    const age = parseInt(record.age);
    const dob = new Date(record.date_of_birth);

    if (isNaN(age) || !Number.isInteger(age)) {
      errors.push(`Age is invalid in Age record  at id ${i + 1} and line number ${i + 1} ,Age should be valid number eg:45`);
      continue;
    }

    if (isNaN(dob.getTime())) {
      errors.push(`Date of birth is in invalid format at date_of_birth record  at id ${i + 1} and line number ${i + 1} ,date_of_birth should be valid format eg:2023-05-10`);
      continue;
    }

    if (address.length < 25) {
      errors.push(`Address is short in adress record  id ${i + 1} and line number ${i + 1}, address must have minimum 25 characters`);
      continue;
    }

    validRecords.push([id, name, address, age, dob]);
  }

  if (errors.length > 0) {
    res.status(400).json({ errors });
  } else {
    const insertQuery = 'INSERT INTO details2_tb (Id, name, address, age, date_of_birth) VALUES ?';
    connection.query(insertQuery, [validRecords], (err, result) => {
      if (err) {
        console.error('Error inserting records:', err);
        res.status(500).json({ error: 'An error occurred while inserting records into the database' });
      } else {
        res.status(200).json({ message: 'Records inserted successfully' });
      }
    });
  }
});

app.listen(3000, () => {
  console.log('Server is running on port 3000');
});
