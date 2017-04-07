//! Spreadsheet file handling and processing.
use formats;
use Result;
use rusqlite::Connection;
use std::cell::Cell;
use std::cmp::{max, min};
use std::path::Path;


/// A loaded spreadsheet file. Provides methods for loading, saving, reading, and editing.
pub struct Spreadsheet {
    /// Open SQLite database for storing spreadsheet data.
    database: Connection,
    /// Whether the spreadhseet has been modified.
    dirty: Cell<bool>,
    /// Number of rows in the spreadsheet.
    row_count: Cell<i64>,
}

/// Position for inserting values at.
#[derive(Clone, Copy, Eq, PartialEq)]
pub enum InsertPosition {
    Index(i64),
    End,
}

impl Spreadsheet {
    /// Create a new, blank spreadsheet.
    pub fn new() -> Self {
        // Open an on-disk, temporary scratch database.
        let connection = Connection::open("").unwrap();

        // Set up the schema.
        connection.execute_batch("
            CREATE TABLE columns (
                id          INTEGER PRIMARY KEY NOT NULL,
                name        TEXT NOT NULL
            );

            CREATE TABLE cells (
                column      INTEGER NOT NULL,
                row         INTEGER NOT NULL,
                value       TEXT
            );
        ").unwrap();

        Self {
            database: connection,
            dirty: Cell::new(false),
            row_count: Cell::new(0),
        }
    }

    /// Open a spreadsheet from a file.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let path = path.as_ref();

        let loader = match path.extension().and_then(|s| s.to_str()) {
            Some("csv") => formats::load_csv,
            _ => return Err("Unknown file extension.".into()),
        };

        let spreadsheet = Self::new();
        loader(path, &spreadsheet)?;
        spreadsheet.clear_dirty();

        Ok(spreadsheet)
    }

    /// Save the spreadsheet to a file.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        formats::save_csv(path.as_ref(), self)?;
        self.clear_dirty();

        Ok(())
    }

    /// Check if the spreadsheet has been modified.
    pub fn is_dirty(&self) -> bool {
        self.dirty.get()
    }

    /// Clear the dirty flag if set.
    pub fn clear_dirty(&self) {
        self.dirty.set(false);
    }

    /// Get the number of columns in the spreadsheet.
    pub fn get_column_count(&self) -> i64 {
        self.database.query_row("SELECT COUNT(id) FROM columns", &[], |row| {
            row.get(0)
        }).unwrap()
    }

    /// Get the spreadsheet columns.
    pub fn get_columns(&self) -> Vec<String> {
        let mut stmt = self.database.prepare_cached("SELECT name FROM columns ORDER BY id ASC").unwrap();
        let mut rows = stmt.query(&[]).unwrap();
        let mut columns = Vec::new();

        while let Some(Ok(row)) = rows.next() {
            columns.push(row.get(0));
        }

        columns
    }

    /// Inserts columns starting at the given position.
    pub fn insert_columns(&self, position: InsertPosition, names: Vec<String>) -> Result<()> {
        // Get the absolute index to insert at.
        let position = match position {
            InsertPosition::Index(i) => i,
            InsertPosition::End => self.get_column_count(),
        };

        // Shift columns to the right to make room for the given column count.
        let shift_amount = names.len() as i64;
        self.database.execute("
            UPDATE columns
            SET id = id + ?
            WHERE id >= ?
        ", &[
            &shift_amount,
            &position,
        ])?;

        // Insert the new columns.
        let mut stmt = self.database.prepare_cached("INSERT INTO columns (id, name) VALUES (?, ?)")?;
        for (offset, value) in names.into_iter().enumerate() {
            let pos = position + offset as i64;
            stmt.execute(&[&pos, &value])?;
        }

        self.dirty.set(true);

        Ok(())
    }

    /// Get the value of a specific cell.
    pub fn get_cell(&self, row: i64, column: i64) -> Option<String> {
        self.database.query_row("
            SELECT value FROM cells
            WHERE row = ? AND column = ?
        ", &[&row, &column], |row| {
            row.get(0)
        }).unwrap_or(None)
    }

    /// Set the value of a specific cell.
    pub fn set_cell<S: Into<Option<String>>>(&self, row: i64, column: i64, value: S) -> Result<()> {
        self.database.execute("
            UPDATE cells
            SET value = ?
            WHERE row = ? AND column = ?
        ", &[
            &value.into(),
            &row,
            &column
        ])?;

        self.dirty.set(true);

        Ok(())
    }

    /// Get the number of rows in the spreadsheet.
    pub fn get_row_count(&self) -> i64 {
        self.row_count.get()
    }

    /// Get a range of values.
    pub fn get_rows(&self, start: i64, end: i64) -> Result<Vec<Vec<Option<String>>>> {
        let mut stmt = self.database.prepare("
            SELECT row, value FROM cells
            WHERE row >= ? AND row <= ?
            ORDER BY row, column ASC
        ")?;

        info!("loading spreadsheet values in rows {} - {}", start, end);
        let mut results = stmt.query(&[&start, &end])?;
        let mut rows = Vec::new();
        let mut row = Vec::new();

        while let Some(result) = results.next() {
            let result = result?;
            let current_row: i64 = result.get(0);

            if current_row - start > rows.len() as i64 {
                rows.push(row);
                row = Vec::new();
            }

            row.push(result.get(1));
        }

        rows.push(row);
        info!("got back {} rows", rows.len());

        Ok(rows)
    }

    /// Insert a row into the spreadsheet beginning at the specified position.
    pub fn insert_row(&self, position: InsertPosition, values: Vec<String>) -> Result<()> {
        // Get the absolute index to insert at.
        let row = match position {
            InsertPosition::Index(i) => i,
            InsertPosition::End => self.get_row_count(),
        };

        // Shift rows below down by one.
        if position != InsertPosition::End {
            self.database.execute("
                UPDATE cells
                SET row = row + 1
                WHERE row >= ?
            ", &[&row])?;
        }

        // Insert the cells into the new row.
        let mut cell_stmt = self.database.prepare_cached("INSERT INTO cells (column, row, value) VALUES (?, ?, ?)")?;
        for (pos, value) in values.into_iter().enumerate() {
            let column = pos as i64;
            cell_stmt.execute(&[&column, &row, &value])?;
        }

        self.row_count.set(self.row_count.get() + 1);
        self.dirty.set(true);

        Ok(())
    }

    /// Delete a range of rows.
    pub fn delete_rows(&self, start: i64, end: i64) -> Result<()> {
        let start = max(0, min(self.get_row_count(), start));
        let end = max(0, min(self.get_row_count(), end));

        if start > end {
            return Err("Starting row must be greater than or equal to the ending row".into());
        }

        let count = end - start + 1;
        info!("deleting {} rows ({} - {})", count, start, end);

        // Delete cells belonging to the rows.
        self.database.execute("
            DELETE FROM cells
            WHERE row >= ? AND row <= ?
        ", &[
            &start,
            &end,
        ])?;

        // Shift rows after the deleted range back up.
        self.database.execute("
            UPDATE cells
            SET row = row - ?
            WHERE row > ?
        ", &[
            &count,
            &end,
        ])?;

        self.row_count.set(self.get_row_count() - count);
        self.dirty.set(true);

        Ok(())
    }
}
