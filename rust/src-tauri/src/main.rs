// Скрывает окно консоли при запуске на Windows
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    otkryvashka_lib::run()
}
