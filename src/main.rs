extern crate csv;
extern crate gio;
extern crate gtk;
#[macro_use]
extern crate log;
extern crate rusqlite;
extern crate simplelog;

mod spreadsheet;
mod ui;

use gtk::prelude::*;
use gtk::Application;

pub static LICENSE: &'static str = include_str!("../LICENSE");
pub static VERSION: &'static str = env!("CARGO_PKG_VERSION");


fn startup(application: &Application) {
    // Build the application menu.
    let app_menu = ui::build_app_menu();
    application.set_app_menu(Some(&app_menu));

    // Build the menu bar.
    let window_menu = ui::build_window_menu();
    application.set_menubar(Some(&window_menu));

    let quit_action = gio::SimpleAction::new("quit", None);
    let cloned = application.clone();
    quit_action.connect_activate(move |_,_| {
        quit(&cloned);
    });
    application.add_action(&quit_action);

    let about_action = gio::SimpleAction::new("about", None);
    about_action.connect_activate(move |_,_| {
        ui::show_about_dialog();
    });
    application.add_action(&about_action);

    // Create the main window.
    let window = ui::MainWindow::new(&application);
    window.window().show_all();
}

fn activate(application: &Application) {}

fn quit(application: &Application) {
    info!("quit");
    for window in application.get_windows() {
        window.close();
    }
}

fn main() {
    let _ = simplelog::TermLogger::init(log::LogLevelFilter::Info, simplelog::Config::default());

    let application = Application::new(
        Some("com.widen.astinus"),
        gio::APPLICATION_HANDLES_OPEN)
    .expect("failed to initialize application");

    application.connect_startup(startup);
    application.connect_activate(activate);

    application.run(0, &[]);
}
