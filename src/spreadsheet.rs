//! Spreadsheet file handling and processing.
use csv;
use rusqlite::Connection;
use std::error::Error;
use std::path::Path;


pub struct Spreadsheet {
    database: Connection,
}

impl Spreadsheet {
    /// Create a new, blank spreadsheet.
    pub fn new() -> Self {
        // Open an on-disk, temporary database.
        let connection = Connection::open("").unwrap();

        // Set up the schema.
        connection.execute_batch("
            CREATE TABLE columns (
                id          INTEGER PRIMARY KEY NOT NULL,
                name        TEXT NOT NULL
            );

            CREATE TABLE cells (
                row         INTEGER NOT NULL,
                column      INTEGER NOT NULL,
                value       TEXT,
                PRIMARY KEY (row, column),
                FOREIGN KEY (column) REFERENCES columns(id)
            );
        ").unwrap();

        Self {
            database: connection,
        }
    }

    /// Open a spreadsheet from a file.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self, Box<Error>> {
        let path = path.as_ref();

        match path.extension().and_then(|s| s.to_str()) {
            Some("csv") => Self::open_csv(path),
            _ => Err("Unknown file extension.".into()),
        }
    }

    /// Read a CSV file into a spreadsheet.
    fn open_csv(path: &Path) -> Result<Self, Box<Error>> {
        let mut reader = csv::Reader::from_file(path)?;
        let spreadsheet = Self::new();

        // Load the headers from the CSV first.
        {
            let mut stmt = spreadsheet.database.prepare("INSERT INTO columns (id, name) VALUES (?, ?)")?;

            for (pos, header) in reader.headers()?.iter().enumerate() {
                stmt.execute(&[&(pos as u32), header])?;
            }
        }

        // Read all rows in the file and insert them into the database.
        {
            let mut stmt = spreadsheet.database.prepare("INSERT INTO cells (row, column, value) VALUES (?, ?, ?)")?;
            let mut records = reader.records();
            let mut row = 0;

            while let Some(Ok(record)) = records.next() {
                for (pos, value) in record.iter().enumerate() {
                    stmt.execute(&[&row, &(pos as u32), value])?;
                }

                row += 1;
            }
        }

        Ok(spreadsheet)
    }

    /// Get the number of columns in the spreadsheet.
    pub fn width(&self) -> u32 {
        self.database.query_row("SELECT COUNT(id) FROM columns", &[], |row| {
            row.get(0)
        }).unwrap()
    }

    /// Get the spreadsheet columns.
    pub fn get_columns(&self) -> Vec<String> {
        let mut stmt = self.database.prepare("SELECT name FROM columns ORDER BY id ASC").unwrap();
        let mut rows = stmt.query(&[]).unwrap();
        let mut columns = Vec::new();

        while let Some(Ok(row)) = rows.next() {
            columns.push(row.get(0));
        }

        columns
    }

    /// Get a range of values.
    pub fn get_range(&self, start: u64, end: u64) -> Result<Vec<Vec<Option<String>>>, Box<Error>> {
        let mut stmt = self.database.prepare("
            SELECT row, value FROM cells
            WHERE row >= ? AND row <= ?
            ORDER BY row, column ASC
        ")?;

        info!("loading spreadsheet values in range {} - {}", start, end);
        let mut results = stmt.query(&[&(start as i64), &(end as i64)])?;
        let mut rows = Vec::new();
        let mut row = Vec::new();

        while let Some(Ok(result)) = results.next() {
            let current_row_offset: i64 = result.get(0);

            if current_row_offset > rows.len() as i64 {
                rows.push(row);
                row = Vec::new();
            }

            row.push(result.get(1));
        }

        Ok(rows)
    }
}
