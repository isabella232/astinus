use rusqlite::Connection;


/// Load a CSV file into a database.
fn load_csv(path: &Path, database: &Connection) -> Result<()), Box<Error>> {
    let mut reader = csv::Reader::from_file(path)?;

    // Load the headers from the CSV first.
    {
        let mut stmt = database.prepare("INSERT INTO columns (id, name) VALUES (?, ?)")?;

        for (pos, header) in reader.headers()?.iter().enumerate() {
            stmt.execute(&[&(pos as u32), header])?;
        }
    }

    // Read all rows in the file and insert them into the database.
    {
        let mut stmt = database.prepare("INSERT INTO cells (row, column, value) VALUES (?, ?, ?)")?;
        let mut records = reader.records();
        let mut row = 0;

        while let Some(Ok(record)) = records.next() {
            for (pos, value) in record.iter().enumerate() {
                stmt.execute(&[&row, &(pos as u32), value])?;
            }

            row += 1;
        }
    }

    Ok(()))
}
