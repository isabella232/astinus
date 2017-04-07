use gio::{Menu, SimpleAction};
use gio::prelude::*;
use gtk::*;
use Result;
use spreadsheet::Spreadsheet;
use std::boxed::Box;
use std::cell::{Cell, RefCell};
use std::cmp::{max, min};
use std::error::Error;
use std::path::*;
use std::rc::Rc;


const PAGE_SIZE: i64 = 1000;


pub fn build_app_menu() -> Menu {
    let menu = Menu::new();

    menu.append("About Astinus", "app.about");
    menu.append("Quit", "app.quit");

    menu
}

pub fn build_window_menu() -> Menu {
    let menu = Menu::new();

    let file_menu = Menu::new();
    file_menu.append("New File", "win.new");
    file_menu.append("Open...", "win.open");
    file_menu.append("Save As...", "win.save");
    file_menu.append("Close", "win.close");
    menu.append_submenu("File", &file_menu);

    menu
}

pub fn show_about_dialog() {
    let dialog = AboutDialog::new();

    dialog.set_authors(&["Stephen Coakley <scoakley@widen.com>"]);
    dialog.set_copyright("Copyright (c) 2017 Widen Enterprises, Inc");
    dialog.set_license(::LICENSE);
    dialog.set_program_name("Astinus");
    dialog.set_version(::VERSION);

    dialog.run();
    dialog.destroy();
}


#[derive(Clone)]
pub struct MainWindow {
    builder: Builder,
    window: ApplicationWindow,
    page_entry: SpinButton,
    spreadsheet_view: TreeView,
    status_bar: Statusbar,
    delete_dialog: Dialog,
    open_dialog: FileChooserDialog,
    save_dialog: FileChooserDialog,
    spreadsheet: Rc<RefCell<Option<Spreadsheet>>>,
    page: Rc<Cell<i64>>,
}

impl MainWindow {
    pub fn new(application: &Application) -> Self {
        let builder = Builder::new();
        builder.add_from_string(include_str!("main.glade")).unwrap();

        let main = Self {
            builder: builder.clone(),
            window: builder.get_object("window").unwrap(),
            page_entry: builder.get_object("page_entry").unwrap(),
            spreadsheet_view: builder.get_object("spreadsheet_view").unwrap(),
            status_bar: builder.get_object("status_bar").unwrap(),
            delete_dialog: builder.get_object("delete_dialog").unwrap(),
            open_dialog: builder.get_object("open_dialog").unwrap(),
            save_dialog: builder.get_object("save_dialog").unwrap(),
            spreadsheet: Rc::new(RefCell::new(None)),
            page: Rc::new(Cell::new(1)),
        };

        let window: ApplicationWindow = builder.get_object("window").unwrap();
        window.set_application(Some(application));

        window.add_action(&create_action("open", &main, true, |main| {
            main.show_open_dialog();
        }));

        window.add_action(&create_action("save", &main, false, |main| {
            main.show_save_dialog();
        }));

        window.add_action(&create_action("close", &main, false, |main| {
            main.close_file();
        }));

        window.add_action(&create_action("previous_page", &main, false, |main| {
            main.go_to_previous_page();
        }));

        window.add_action(&create_action("next_page", &main, false, |main| {
            main.go_to_next_page();
        }));

        window.add_action(&create_action("delete", &main, false, |main| {
            main.show_delete_dialog();
        }));

        {
            let cloned = main.clone();
            window.connect_delete_event(move |_, _| {
                cloned.close_file();
                Inhibit(false)
            });
        }

        {
            let cloned = main.clone();
            main.page_entry.connect_value_changed(move |page_entry| {
                cloned.go_to_page(page_entry.get_value_as_int() as i64);
            });
        }

        main.update_state();

        main
    }

    /// Get the main window object.
    pub fn window(&self) -> ApplicationWindow {
        self.window.clone()
    }

    /// Check if a file is currently opened in this window.
    pub fn is_file_opened(&self) -> bool {
        self.spreadsheet.borrow().is_some()
    }

    pub fn open_file<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.close_file();

        let spreadsheet = Spreadsheet::open(path)?;
        *self.spreadsheet.borrow_mut() = Some(spreadsheet);

        self.page.set(1);
        self.prepare_spreadsheet_view();
        self.update_spreadsheet_view()?;
        self.update_state();

        Ok(())
    }

    /// Save the active file if one is open.
    pub fn save_file<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        if let Some(spreadsheet) = self.spreadsheet.borrow().as_ref() {
            spreadsheet.save(path)?;
        }

        Ok(())
    }

    /// Close the active file if one is open.
    pub fn close_file(&self) {
        if let Some(spreadsheet) = self.spreadsheet.borrow_mut().take() {
            if spreadsheet.is_dirty() {
                let dialog = MessageDialog::new(
                    Some(&self.window),
                    DIALOG_MODAL,
                    MessageType::Warning,
                    ButtonsType::YesNo,
                    "The current spreadsheet has not been saved. Would you like to save it?"
                );

                dialog.set_modal(true);
                dialog.set_position(WindowPosition::CenterOnParent);
                dialog.set_resizable(false);
                dialog.set_urgency_hint(true);
                let response = dialog.run();
                dialog.destroy();

                if response == ResponseType::Yes.into() {
                    self.show_save_dialog();
                }
            }
        }

        self.prepare_spreadsheet_view();
        self.update_state();
    }

    /// Get the current page being viewed.
    pub fn get_current_page(&self) -> i64 {
        self.page.get()
    }

    /// Get the total number of pages.
    pub fn get_page_count(&self) -> i64 {
        if let Some(spreadsheet) = self.spreadsheet.borrow().as_ref() {
            spreadsheet.get_row_count() / PAGE_SIZE + 1
        } else {
            0
        }
    }

    /// Get the total number of rows.
    pub fn get_row_count(&self) -> i64 {
        if let Some(spreadsheet) = self.spreadsheet.borrow().as_ref() {
            spreadsheet.get_row_count()
        } else {
            0
        }
    }

    /// Get the first row number currently being displayed.
    pub fn get_first_row_offset(&self) -> i64 {
        (self.get_current_page() - 1) * PAGE_SIZE
    }

    /// Get the last row number currently being displayed.
    pub fn get_last_row_offset(&self) -> i64 {
        min(
            self.get_first_row_offset() + PAGE_SIZE - 1,
            self.get_row_count() - 1,
        )
    }

    /// Jump the spreadsheet view to a page.
    pub fn go_to_page(&self, page: i64) {
        let page = max(1, min(page, self.get_page_count()));

        if self.get_current_page() != page {
            self.page.set(page);

            self.update_spreadsheet_view()
                .unwrap_or_else(|e| self.show_error_dialog(e));
            self.update_state();
        }
    }

    /// Go to the next page of spreadsheet rows.
    pub fn go_to_next_page(&self) {
        self.go_to_page(self.get_current_page() + 1);
    }

    /// Go to the previous page of spreadsheet rows.
    pub fn go_to_previous_page(&self) {
        self.go_to_page(self.get_current_page() - 1);
    }

    pub fn show_delete_dialog(&self) {
        if self.delete_dialog.run() == ResponseType::Ok.into() {
            let from_entry: Entry = self.builder.get_object("delete_rows_from_entry").unwrap();
            let to_entry: Entry = self.builder.get_object("delete_rows_to_entry").unwrap();

            let from: String = from_entry.get_text().unwrap();
            let to: String = to_entry.get_text().unwrap();
            info!("from: {} to: {}", from, to);
            let from: i64 = from.parse().unwrap();
            let to: i64 = to.parse().unwrap();

            if from <= to {
                if let Some(spreadsheet) = self.spreadsheet.borrow().as_ref() {
                    spreadsheet.delete_rows(from - 1, to - 1)
                        .unwrap_or_else(|e| self.show_error_dialog(e));
                }

                self.update_spreadsheet_view()
                    .unwrap_or_else(|e| self.show_error_dialog(e));
                self.update_state();
            } else {
                warn!("starting row to delete is larger then ending row: {} > {}", from, to);
            }
        }

        self.delete_dialog.hide();
    }

    pub fn show_open_dialog(&self) {
        let mut filename = None;

        if self.open_dialog.get_filter().is_none() {
            let text_filter = FileFilter::new();
            text_filter.set_name("Text files");
            text_filter.add_pattern("*.csv");
            text_filter.add_pattern("*.tsv");
            text_filter.add_pattern("*.txt");
            self.open_dialog.add_filter(&text_filter);

            let excel_filter = FileFilter::new();
            excel_filter.set_name("Excel spreadsheet");
            excel_filter.add_pattern("*.xls");
            excel_filter.add_pattern("*.xlsx");
            self.open_dialog.add_filter(&excel_filter);
        }

        if self.open_dialog.run() == ResponseType::Ok.into() {
            filename = self.open_dialog.get_filename();
        }
        self.open_dialog.hide();

        if let Some(filename) = filename {
            self.open_file(filename)
                .unwrap_or_else(|e| self.show_error_dialog(e));
        }
    }

    pub fn show_save_dialog(&self) {
        let mut filename = None;

        if self.save_dialog.run() == ResponseType::Ok.into() {
            filename = self.save_dialog.get_filename();
        }
        self.save_dialog.hide();

        if let Some(filename) = filename {
            self.save_file(filename)
                .unwrap_or_else(|e| self.show_error_dialog(e));
        }
    }

    fn show_error_dialog(&self, error: Box<Error>) {
        error!("Error: {:?}", error);
        let message = format!("Error: {:?}", error);

        let window = self.window();
        let dialog = MessageDialog::new(
            Some(&window),
            DIALOG_MODAL,
            MessageType::Error,
            ButtonsType::Ok,
            &message
        );

        dialog.set_modal(true);
        dialog.set_position(WindowPosition::CenterOnParent);
        dialog.set_urgency_hint(true);
        dialog.run();
        dialog.destroy();
    }

    /// Update the UI based on the current state of the window.
    fn update_state(&self) {
        // Update window actions.
        let file_actions = self.is_file_opened();
        self.set_action_enabled("save", file_actions);
        self.set_action_enabled("close", file_actions);
        self.set_action_enabled("previous_page", file_actions && self.get_current_page() > 1);
        self.set_action_enabled("next_page", file_actions && self.get_current_page() < self.get_page_count());
        self.set_action_enabled("delete", file_actions);

        // Update the page entry.
        self.page_entry.set_range(1.0, self.get_page_count() as f64);
        self.page_entry.set_value(self.get_current_page() as f64);

        // Update the status bar contents.
        let page_status = format!(
            "Page {} of {} (rows {} - {}) of {} rows",
            self.get_current_page(),
            self.get_page_count(),
            self.get_first_row_offset() + 1,
            self.get_last_row_offset() + 1,
            self.get_row_count(),
        );
        self.status_bar.remove_all(0);
        self.status_bar.push(0, &page_status);

        // Update the titlebar.
        if let Some(spreadsheet) = self.spreadsheet.borrow().as_ref() {
            let title = format!("Astinus - {}", spreadsheet.name());
            self.window.set_title(&title);
        } else {
            self.window.set_title("Astinus");
        }
    }

    /// Enable or disable a window action.
    fn set_action_enabled(&self, action: &str, enabled: bool) {
        let action: SimpleAction = self.window().lookup_action(action).unwrap().downcast().unwrap();
        action.set_enabled(enabled);
    }

    /// Prepare the spreadsheet view for displaying the current spreadsheet.
    fn prepare_spreadsheet_view(&self) {
        let spreadsheet = self.spreadsheet.borrow();

        // Remove the previous model.
        self.spreadsheet_view.set_model::<TreeModel>(None);

        // Remove all existing columns.
        for column in self.spreadsheet_view.get_columns() {
            self.spreadsheet_view.remove_column(&column);
        }

        if let Some(spreadsheet) = spreadsheet.as_ref() {
            // Populate new columns.
            for (index, title) in spreadsheet.get_columns().into_iter().enumerate() {
                let column = TreeViewColumn::new();
                column.set_resizable(true);
                column.set_title(&title);

                let renderer = CellRendererText::new();
                renderer.set_property_editable(true);
                let cloned = self.clone();
                renderer.connect_edited(move |r, p, v| {
                    cloned.on_edit(r, index as i64, p, v)
                });

                column.pack_start(&renderer, true);
                column.add_attribute(&renderer, "text", index as i32);

                self.spreadsheet_view.append_column(&column);
            }

            // Create a new model.
            let mut column_types = Vec::new();
            for _ in 0..spreadsheet.get_column_count() {
                column_types.push(Type::String);
            }

            let model = ListStore::new(&column_types);
            self.spreadsheet_view.set_model(Some(&model));
        }
    }

    /// Update the current spreadsheet view.
    fn update_spreadsheet_view(&self) -> Result<()> {
        // Compute start and end ranges.
        let start = self.get_first_row_offset();
        let end = self.get_last_row_offset();

        let spreadsheet = self.spreadsheet.borrow();

        if let Some(spreadsheet) = spreadsheet.as_ref() {
            if let Some(model) = self.spreadsheet_view.get_model() {
                let model: ListStore = model.downcast().unwrap();

                model.clear();

                for row in spreadsheet.get_rows(start, end)? {
                    let iter = model.append();

                    for (column, cell) in row.into_iter().enumerate() {
                        let value = cell.as_ref().into();
                        model.set_value(&iter, column as u32, &value);
                    }
                }

                self.spreadsheet_view.set_model(Some(&model));
            }
        }

        Ok(())
    }

    fn on_edit(&self, _: &CellRendererText, column: i64, path: TreePath, value: &str) {
        let spreadsheet = self.spreadsheet.borrow();
        let row_offset = path.get_indices()[0] as i64;
        let row = self.get_first_row_offset() + row_offset;

        if let Some(spreadsheet) = spreadsheet.as_ref() {
            spreadsheet.set_cell(row as i64, column, Some(value.to_string()))
                .unwrap_or_else(|e| self.show_error_dialog(e));
        }

        if let Some(model) = self.spreadsheet_view.get_model() {
            let model: ListStore = model.downcast().unwrap();
            let iter = model.get_iter(&path).unwrap();

            let value = value.to_value();
            model.set_value(&iter, column as u32, &value);
        }
    }
}

/// Create an action mapping.
pub fn create_action<T, F>(name: &str, context: &T, enabled: bool, f: F) -> SimpleAction
    where T: Clone + 'static, F: Fn(T) + 'static
{
    let action = SimpleAction::new(name, None);
    action.set_enabled(enabled);

    let context = context.clone();
    action.connect_activate(move |_, _| {
        let cloned = context.clone();
        f(cloned);
    });

    action
}
