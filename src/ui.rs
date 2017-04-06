use gio::Menu;
use gtk::*;
use gtk::prelude::*;
use spreadsheet::Spreadsheet;
use std::boxed::Box;
use std::cell::RefCell;
use std::path::*;
use std::rc::Rc;
use std::error::Error;


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
    spreadsheet: Rc<RefCell<Option<Spreadsheet>>>,
}

impl MainWindow {
    pub fn new(application: &Application) -> Self {
        let builder = Builder::new();
        builder.add_from_string(include_str!("main.glade")).unwrap();

        let main = Self {
            builder: builder.clone(),
            spreadsheet: Rc::new(RefCell::new(None)),
        };

        let window: Window = builder.get_object("window").unwrap();
        window.set_application(Some(application));

        let open_button: Button = builder.get_object("open_button").unwrap();
        let cloned = main.clone();
        open_button.connect_clicked(move |_| {
            cloned.show_open_dialog();
        });

        main
    }

    /// Get the main window object.
    pub fn window(&self) -> Window {
        self.builder.get_object("window").unwrap()
    }

    pub fn open_file<P: AsRef<Path>>(&self, path: P) {
        match Spreadsheet::open(path) {
            Ok(spreadsheet) => {
                self.prepare_spreadsheet_view(&spreadsheet);
                *self.spreadsheet.borrow_mut() = Some(spreadsheet);

                if let Err(e) = self.set_spreadsheet_view(0, 1000) {
                    self.show_error_dialog(e);
                }
            }
            Err(e) => {
                self.show_error_dialog(e);
            }
        }
    }

    fn show_error_dialog(&self, error: Box<Error>) {
        error!("Error: {:?}", error);
        let message = format!("Error: {:?}", error);

        let window = self.window();
        let dialog = MessageDialog::new(
            Some(&window),
            DialogFlags::empty(),
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

    fn show_open_dialog(&self) {
        let window = self.window();
        let file_chooser = FileChooserDialog::new(Some("Open spreadsheet"), Some(&window), FileChooserAction::Open);
        let mut filename = None;

        file_chooser.add_buttons(&[
            ("Open", ResponseType::Ok.into()),
            ("Cancel", ResponseType::Cancel.into()),
        ]);
        file_chooser.set_position(WindowPosition::CenterOnParent);

        let text_filter = FileFilter::new();
        text_filter.set_name("Text files");
        text_filter.add_pattern("*.csv");
        text_filter.add_pattern("*.tsv");
        text_filter.add_pattern("*.txt");
        file_chooser.add_filter(&text_filter);

        let excel_filter = FileFilter::new();
        excel_filter.set_name("Excel spreadsheet");
        excel_filter.add_pattern("*.xls");
        excel_filter.add_pattern("*.xlsx");
        file_chooser.add_filter(&excel_filter);

        if file_chooser.run() == ResponseType::Ok.into() {
            filename = file_chooser.get_filename();
        }
        file_chooser.destroy();

        if let Some(filename) = filename {
            self.open_file(filename);
        }
    }

    /// Prepare the spreadsheet view for displaying the given spreadsheet.
    fn prepare_spreadsheet_view(&self, spreadsheet: &Spreadsheet) {
        let spreadsheet_view: TreeView = self.builder.get_object("spreadsheet_view").unwrap();

        // Remove the previous model.
        spreadsheet_view.set_model::<TreeModel>(None);

        // Remove all existing columns.
        for column in spreadsheet_view.get_columns() {
            spreadsheet_view.remove_column(&column);
        }

        // Populate new columns.
        for (index, title) in spreadsheet.get_columns().into_iter().enumerate() {
            let column = TreeViewColumn::new();
            column.set_resizable(true);
            column.set_title(&title);

            let renderer = CellRendererText::new();
            column.pack_start(&renderer, true);
            column.add_attribute(&renderer, "text", index as i32);

            spreadsheet_view.append_column(&column);
        }

        // Create a new model.
        let mut column_types = Vec::new();
        for _ in 0..spreadsheet.get_column_count() {
            column_types.push(Type::String);
        }

        let model = ListStore::new(&column_types);
        spreadsheet_view.set_model(Some(&model));
    }

    /// Set the current spreadsheet view.
    fn set_spreadsheet_view(&self, start: i64, end: i64) -> Result<(), Box<Error>> {
        let spreadsheet_view: TreeView = self.builder.get_object("spreadsheet_view").unwrap();
        let spreadsheet = self.spreadsheet.borrow();

        if let Some(spreadsheet) = spreadsheet.as_ref() {
            if let Some(model) = spreadsheet_view.get_model() {
                let model: ListStore = model.downcast().unwrap();

                model.clear();

                for row in spreadsheet.get_rows(start, end)? {
                    let iter = model.append();

                    for (column, cell) in row.into_iter().enumerate() {
                        let value = cell.as_ref().into();
                        model.set_value(&iter, column as u32, &value);
                    }
                }

                spreadsheet_view.set_model(Some(&model));
            }
        }

        Ok(())
    }

    // fn create_list_store(spreadsheet: &Spreadsheet) -> ListStore {}
}
