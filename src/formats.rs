use csv;
use Result;
use spreadsheet::*;
use std::path::Path;


/// Load a CSV file into a database.
pub fn load_csv(path: &Path, spreadsheet: &Spreadsheet) -> Result<()> {
    let mut reader = csv::Reader::from_file(path)?;

    // Load the headers from the CSV first.
    spreadsheet.insert_columns(InsertPosition::End, reader.headers()?)?;

    // Read all rows in the file and insert them into the database.
    let mut records = reader.records();
    while let Some(record) = records.next() {
        spreadsheet.insert_row(InsertPosition::End, record?)?;
    }

    Ok(())
}

/// Save a spreadsheet to a CSV file.
pub fn save_csv(path: &Path, spreadsheet: &Spreadsheet) -> Result<()> {
    let mut writer = csv::Writer::from_file(path)?;

    writer.encode(spreadsheet.get_columns())?;

    for row in spreadsheet.get_rows(0, spreadsheet.get_row_count() - 1) {
        writer.encode(row)?;
    }

    Ok(())
}
